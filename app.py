# 入居管理表アプリ（改修版）
# - 西澤様の追加要件＆入居管理表サンプル（v2）準拠
# - 5行ブロック（家賃/共益費/駐車料/水道料/合計）
# - Pxx(駐車場)行は備考の(0001)等を見て対象室の駐車料へ自動付替え
# - 同一室で入退去があれば賃借人ごとにブロックを分ける
# - Excel は合計欄に SUM, 最下段集計に SUMIF を使用
# - 礼金・更新料は右端の結合セルに契約単位で合算表示、水道料欄追加

import streamlit as st
import io
import json
import asyncio
import base64
import re
import logging
from datetime import datetime
from pathlib import Path
from PIL import Image
import fitz  # PyMuPDF
import pdfplumber
from openai import AsyncOpenAI
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ===== ログ =====
if not logging.getLogger().hasHandlers():
    log_filename = f"app_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format='[%(asctime)s] %(levelname)s: %(message)s',
        handlers=[logging.FileHandler(log_filename, encoding='utf-8'),
                  logging.StreamHandler()]
    )
logger = logging.getLogger(__name__)

# ===== OpenAI 非同期クライアント =====
client = AsyncOpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ========== Vision: PDF → JSON 抽出 ==========
VISION_INSTRUCTIONS = (
    "あなたは不動産管理アシスタントです。収支報告書（送金明細書）から、"
    "各『室番号×賃借人（契約）』の入居情報を抽出し、厳格な JSON で返してください。\n"
    "要件:\n"
    "1) 出力は必ず次のトップレベル構造:\n"
    "{\n"
    "  \"records\": [\n"
    "    {\n"
    "      \"room\": \"0101\" または \"P01\" など,\n"
    "      \"tenant\": \"賃借人名\"（駐車場(Pxx)は空文字でも可）, \n"
    "      \"monthly\": {\n"
    "        \"YYYY-MM\": {\n"
    "          \"rent\": 家賃, \"fee\": 共益費, \"parking\": 駐車料, \"water\": 水道料,\n"
    "          \"reikin\": 礼金, \"koushin\": 更新料, \"bikou\": \"備考文字列\"\n"
    "        }, ...\n"
    "      },\n"
    "      \"shikikin\": 敷金合計（分かれば。なければ0）, \n"
    "      \"linked_room\": \"0001\" のように、Pxx行が特定住戸に紐付く場合に記す（備考の（0001）表記等から判断）\n"
    "    }, ...\n"
    "  ]\n"
    "}\n"
    "2) 各数値はカンマ無しの整数。空欄は 0。\n"
    "3) 月キーは YYYY-MM（例: 2024-11）。表に現れた全ての月を対象。\n"
    "4) 『P01/P02…』など駐車場の行は必ず room に Pxx を入れ、備考に「（0001）込駐車場」等があれば linked_room に『0001』のように数字4桁で格納。\n"
    "5) 同一室で入退去がある場合は賃借人ごとに別レコード（records の要素を分ける）。\n"
    "6) JSON 以外の文字（前置き・コードブロック）は出力しない。"
)

async def call_openai_vision_async(base64_images, text_context, default_month_id):
    image_parts = [{"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}} for b64 in base64_images]
    messages = [
        {"role": "system", "content": VISION_INSTRUCTIONS},
        {"role": "user", "content": [
            *image_parts,
            {"type": "text", "text":
                f"【OCR補助テキスト】\n{text_context}\n\n"
                f"このPDFには {default_month_id} 付近の月が含まれる可能性があります。"
                f"表内に現れた全ての『年／月』を抽出してください。\n\n"
                "※ 出力は純粋な JSON オブジェクトのみ。"}
        ]}
    ]
    resp = await client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        temperature=0.0,
        max_tokens=4096,
    )
    return resp.choices[0].message.content

# ========== ユーティリティ ==========
def convert_pdf_to_images(pdf_bytes, dpi=220):
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in pdf:
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    return images

def convert_image_to_base64(image):
    buf = io.BytesIO()
    image.save(buf, format="JPEG", quality=90)
    return base64.b64encode(buf.getvalue()).decode("utf-8")

def extract_text_with_pdfplumber(pdf_bytes):
    texts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or "")
    return "\n".join(texts)

def extract_month_from_filename(filename: str) -> str:
    m = re.search(r"(\d{4})年(\d{1,2})月", filename)
    return f"{m.group(1)}-{m.group(2).zfill(2)}" if m else "unknown"

def normalize_room(s: str) -> str:
    """ 0101 / 0205 / 0303 / P01 などへ正規化 """
    if not s:
        return s
    s = str(s).strip()
    if re.fullmatch(r"P\d{1,2}", s, re.IGNORECASE):
        p = s.upper().replace("P", "")
        return f"P{p.zfill(2)}"
    # 数字系は4桁ゼロパディング
    digits = re.sub(r"\D", "", s)
    if digits:
        return digits.zfill(4)
    return s

def clean_int(v):
    if v is None: return 0
    if isinstance(v, (int, float)): return int(v)
    s = str(v).replace(",", "").strip()
    if s == "": return 0
    try: return int(float(s))
    except: return 0

def month_key(s: str) -> str:
    m = re.match(r"(\d{4})[-/年](\d{1,2})", str(s))
    if not m: return s
    return f"{m.group(1)}-{m.group(2).zfill(2)}"

# ========== 1ファイル処理 ==========
async def handle_file(file, max_attempts=3):
    file_name = file.name
    logger.info(f"開始: {file_name}")
    default_month_id = extract_month_from_filename(file_name)
    file_bytes = file.read()
    images = convert_pdf_to_images(file_bytes)
    b64s = [convert_image_to_base64(img) for img in images]
    text_context = extract_text_with_pdfplumber(file_bytes)

    last_err = None
    for attempt in range(1, max_attempts + 1):
        try:
            raw = await call_openai_vision_async(b64s, text_context, default_month_id)
            s = raw.strip()
            # コードフェンスの除去
            s = s.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
            obj = json.loads(s)
            if not isinstance(obj, dict) or "records" not in obj or not isinstance(obj["records"], list):
                raise ValueError("JSON ルートが {'records': [...]} になっていません。")
            # 正規化
            norm_records = []
            for r in obj["records"]:
                room = normalize_room(r.get("room", ""))
                tenant = (r.get("tenant") or "").strip()
                shikikin = clean_int(r.get("shikikin"))
                linked_room = normalize_room(r.get("linked_room", "")) if r.get("linked_room") else ""
                monthly = {}
                for mk, mv in (r.get("monthly") or {}).items():
                    mk2 = month_key(mk)
                    monthly[mk2] = {
                        "rent":      clean_int((mv or {}).get("rent")),
                        "fee":       clean_int((mv or {}).get("fee")),
                        "parking":   clean_int((mv or {}).get("parking")),
                        "water":     clean_int((mv or {}).get("water")),
                        "reikin":    clean_int((mv or {}).get("reikin")),
                        "koushin":   clean_int((mv or {}).get("koushin")),
                        "bikou":     str((mv or {}).get("bikou") or "").strip(),
                    }
                norm_records.append({
                    "room": room, "tenant": tenant, "monthly": monthly,
                    "shikikin": shikikin, "linked_room": linked_room
                })
            logger.info(f"{file_name}: JSON解析成功 / {len(norm_records)}件")
            return norm_records
        except Exception as e:
            last_err = e
            logger.warning(f"{file_name}: JSON解析失敗（{attempt}/{max_attempts}）: {e}")
    st.warning(f"{file_name} の出力がJSONとして解釈できませんでした。")
    logger.error(f"{file_name}: 失敗の最終原因: {last_err}")
    return []

# ========== 全ファイル並列処理 & マージ ==========
def merge_records(all_recs, new_recs):
    """
    all_recs: dict[ (room, tenant) ] -> record
      record = {
        room, tenant, monthly: { 'YYYY-MM': {...} }, shikikin, linked_room
      }
    """
    for r in new_recs:
        key = (r["room"], r["tenant"])
        if key not in all_recs:
            all_recs[key] = {
                "room": r["room"],
                "tenant": r["tenant"],
                "monthly": {},
                "shikikin": clean_int(r.get("shikikin", 0)),
                "linked_room": r.get("linked_room", ""),
            }
        # 敷金は最大（または和でもよいが、ここは最大値採用）
        all_recs[key]["shikikin"] = max(all_recs[key]["shikikin"], clean_int(r.get("shikikin", 0)))
        for mk, mv in (r.get("monthly") or {}).items():
            dst = all_recs[key]["monthly"].setdefault(mk, {
                "rent":0,"fee":0,"parking":0,"water":0,"reikin":0,"koushin":0,"bikou":""
            })
            dst["rent"]    += clean_int(mv.get("rent"))
            dst["fee"]     += clean_int(mv.get("fee"))
            dst["parking"] += clean_int(mv.get("parking"))
            dst["water"]   += clean_int(mv.get("water"))
            dst["reikin"]  += clean_int(mv.get("reikin"))
            dst["koushin"] += clean_int(mv.get("koushin"))
            b = str(mv.get("bikou") or "").strip()
            if b:
                # 既存備考に重複追加しない簡易処理
                if dst["bikou"]:
                    if b not in dst["bikou"]:
                        dst["bikou"] += f", {b}"
                else:
                    dst["bikou"] = b

def fold_parking_Pxx(all_recs):
    """
    Pxx レコードを、linked_room に駐車料として付替える。
    付替え先キーは (linked_room, tenant='') が見つからない場合は
    同室の誰かのレコード（賃借人がいるもの）にまとめる（最初に見つかったもの）。
    """
    to_delete = []
    # 検索用: room -> keys(list)
    by_room = {}
    for key, rec in all_recs.items():
        by_room.setdefault(rec["room"], []).append(key)

    for key, rec in list(all_recs.items()):
        room = rec["room"]
        if not room.upper().startswith("P"):
            continue
        # 付替え先
        target_room = rec.get("linked_room") or ""
        if not target_room:
            # 備考から (dddd) を拾う fallback
            for mk, mv in rec.get("monthly", {}).items():
                m = re.search(r"（?(\d{3,4})）?", mv.get("bikou",""))
                if m:
                    target_room = m.group(1).zfill(4)
                    break
        if not target_room:
            # 付替え不能なら残す（稀ケース）
            logger.info(f"Pxx行 {key} は付替え先不明のため残存")
            continue

        # 候補キー
        target_keys = by_room.get(target_room, [])
        if not target_keys:
            # まだ同室のレコードがない場合、空テナントのレコードを新設
            tkey = (target_room, "")
            all_recs[tkey] = {"room": target_room, "tenant":"", "monthly": {}, "shikikin":0, "linked_room":""}
            by_room.setdefault(target_room, []).append(tkey)
            target_keys = [tkey]

        # 付替えは最初の候補へ
        tkey = target_keys[0]
        target = all_recs[tkey]
        for mk, mv in rec.get("monthly", {}).items():
            dst = target["monthly"].setdefault(mk, {"rent":0,"fee":0,"parking":0,"water":0,"reikin":0,"koushin":0,"bikou":""})
            dst["parking"] += clean_int(mv.get("parking"))
            # 備考に「(P01→0001付替)」をメモ（任意）
            note = f"駐車場({room})→{target_room}"
            if note not in (dst["bikou"] or ""):
                dst["bikou"] = (dst["bikou"] + ", " if dst["bikou"] else "") + note

        to_delete.append(key)

    for key in to_delete:
        all_recs.pop(key, None)

async def process_files(files):
    tasks = [handle_file(file) for file in files]
    results = await asyncio.gather(*tasks)

    # 1) マージ
    all_recs = {}  # key = (room, tenant)
    for recs in results:
        merge_records(all_recs, recs)

    # 2) Pxx 付替え
    fold_parking_Pxx(all_recs)

    # 3) 出力用に並べ替え & 基準額付与
    #    -> list[record] へ
    out = []
    for (room, tenant), rec in all_recs.items():
        # 基準額は各科目の月次最大
        def max_of(k):
            return max([clean_int(v.get(k,0)) for v in rec["monthly"].values()] or [0])

        rec["base_rent"]    = max_of("rent")
        rec["base_fee"]     = max_of("fee")
        rec["base_parking"] = max_of("parking")
        rec["base_water"]   = max_of("water")
        out.append(rec)

    # 室番号数値→名前→月最小 でソート
    def room_sort_key(r):
        rm = r["room"]
        num = 9999
        if rm.upper().startswith("P"):
            num = 9000 + int(re.sub(r"\D","",rm) or 0)  # 駐車は末尾に
        else:
            num = int(re.sub(r"\D","",rm) or 0)
        first_month = sorted(r["monthly"].keys())[0] if r["monthly"] else "9999-99"
        return (num, r["tenant"] or "~", first_month)

    out_sorted = sorted(out, key=room_sort_key)

    # 月リスト（全レコードのユニーク月）
    months = sorted({m for r in out_sorted for m in r["monthly"].keys()})
    return out_sorted, months

# ========== Excel 生成（サンプル準拠） ==========
def combine_bikou_contract(rec):
    """契約全体の備考集合（ユニーク）"""
    s = set()
    for mv in rec.get("monthly", {}).values():
        b = (mv.get("bikou") or "").strip()
        if b: s.add(b)
    return ", ".join(sorted(s))

def export_excel(records, months, property_name):
    wb = Workbook()
    ws = wb.active
    sheet_title = property_name or "入居管理表"
    ws.title = sheet_title

    num_months = len(months)
    # ---- style ----
    header_fill = PatternFill("solid", fgColor="BDD7EE")
    green_fill  = PatternFill("solid", fgColor="CCFFCC")
    gray_fill   = PatternFill("solid", fgColor="DDDDDD")
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    center_vert = Alignment(vertical="center", wrap_text=True)
    bold_font   = Font(bold=True)
    red_font    = Font(color="9C0000")
    number_fmt  = "#,##0"
    thin_border = Border(*[Side(style='thin')] * 4)

    # ---- タイトル & 見出し ----
    # B2: タイトル（B2:W2 程度に結合 / 月数で幅を合わせる）
    last_col_idx = 23  # W=23 （見本と合わせるため固定気味に）
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col_idx)
    if months:
        start_month = months[0].replace("-", "年") + "月"
        end_month   = months[-1].replace("-", "年") + "月"
        ws["B2"] = f"{sheet_title} 入居管理表 （{start_month}〜{end_month}）"
    else:
        ws["B2"] = f"{sheet_title} 入居管理表"
    ws["B2"].font = Font(size=14, bold=True)
    ws["B2"].alignment = center

    # 行5：大見出し
    ws.merge_cells("B5:C5"); ws["B5"] = "賃借人"
    ws.merge_cells("D5:E5"); ws["D5"] = "基準額"
    ws["F5"] = "期首\n未収/前受"
    # 月見出し
    for i, m in enumerate(months):
        mm = int(m[5:])
        ws.cell(row=5, column=7+i, value=f"{mm}月")
    # 右端ラベル
    labels = ["合計", "期末\n未収/前受", "礼金・更新料", "敷金", "備考"]
    for i, lab in enumerate(labels):
        ws.cell(row=5, column=7+num_months+i, value=lab)

    # ヘッダスタイル
    col_bikou = 7 + num_months + 4
    for col in range(2, col_bikou + 1):
        c = ws.cell(row=5, column=col)
        c.fill = header_fill; c.font = bold_font; c.alignment = center

    # ---- データ行（5行ブロック）----
    row = 6
    blocks = []  # (start_row, end_row) for later total rows
    for rec in records:
        room = rec["room"]
        tenant = rec["tenant"]
        base_rent    = rec.get("base_rent",0)
        base_fee     = rec.get("base_fee",0)
        base_parking = rec.get("base_parking",0)
        base_water   = rec.get("base_water",0)
        shikikin     = rec.get("shikikin",0)
        reikin_koushin_total = sum((mv.get("reikin",0)+mv.get("koushin",0)) for mv in rec.get("monthly",{}).values())

        # 左側ラベル
        ws.merge_cells(start_row=row,   start_column=2, end_row=row+4, end_column=2) # B:「室番号」
        ws.cell(row=row, column=2, value="室番号").alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        ws.cell(row=row, column=3, value=room).alignment = center; ws.cell(row=row, column=3).fill = green_fill
        ws.merge_cells(start_row=row+1, start_column=3, end_row=row+4, end_column=3)
        ws.cell(row=row+1, column=3, value=tenant).alignment = center

        # D 列 科目名
        labels = ["家賃","共益費　","駐車料","水道料","合計"]
        for i, lab in enumerate(labels):
            ws.cell(row=row+i, column=4, value=lab)

        # 基準額（E列）
        base_vals = [base_rent, base_fee, base_parking, base_water]
        for i, v in enumerate(base_vals):
            cc = ws.cell(row=row+i, column=5, value=v); cc.number_format = number_fmt
        # 合計行（E）は =SUM(E: の4行）
        ws.cell(row=row+4, column=5, value=None).number_format = number_fmt
        ws.cell(row=row+4, column=5).value = f"=SUM(E{row}:E{row+3})"

        # 期首F列は 0 のまま（必要なら編集可）
        for i in range(5):
            ws.cell(row=row+i, column=6, value=0).number_format = number_fmt
        ws.cell(row=row+4, column=6).value = f"=SUM(F{row}:F{row+3})"

        # 月次（G..）
        rent_sum = fee_sum = park_sum = water_sum = 0
        for i, m in enumerate(months):
            mv = rec.get("monthly", {}).get(m, {})
            vals = [
                clean_int(mv.get("rent")),
                clean_int(mv.get("fee")),
                clean_int(mv.get("parking")),
                clean_int(mv.get("water")),
            ]
            rent_sum  += vals[0]; fee_sum += vals[1]; park_sum += vals[2]; water_sum += vals[3]
            for r_i, v in enumerate(vals):
                cc = ws.cell(row=row+r_i, column=7+i, value=v); cc.number_format = number_fmt
            # 合計行（5行目）は式で
            ws.cell(row=row+4, column=7+i).number_format = number_fmt
            ws.cell(row=row+4, column=7+i).value = f"=SUM({get_column_letter(7+i)}{row}:{get_column_letter(7+i)}{row+3})"

        # S列=各行の合計（G..R）
        col_S = 7 + num_months
        for r_i in range(5):
            cell = ws.cell(row=row+r_i, column=col_S); cell.number_format = number_fmt
            cell.value = f"=SUM({get_column_letter(7)}{row+r_i}:{get_column_letter(6+num_months)}{row+r_i})"

        # T列=期末 未収/前受（0 初期化、合計行は =SUM(T4行)）
        col_T = col_S + 1
        for r_i in range(4):
            ws.cell(row=row+r_i, column=col_T, value=0).number_format = number_fmt
        ws.cell(row=row+4, column=col_T).number_format = number_fmt
        ws.cell(row=row+4, column=col_T).value = f"=SUM(T{row}:T{row+3})"

        # U列=礼金・更新料（5行縦結合）
        col_U = col_T + 1
        ws.merge_cells(start_row=row, start_column=col_U, end_row=row+4, end_column=col_U)
        cu = ws.cell(row=row, column=col_U, value=reikin_koushin_total)
        cu.alignment = center_vert; cu.number_format = number_fmt

        # V列=敷金（5行縦結合）
        col_V = col_U + 1
        ws.merge_cells(start_row=row, start_column=col_V, end_row=row+4, end_column=col_V)
        cv = ws.cell(row=row, column=col_V, value=shikikin)
        cv.alignment = center_vert; cv.number_format = number_fmt

        # W列=備考（5行縦結合）
        col_W = col_V + 1
        ws.merge_cells(start_row=row, start_column=col_W, end_row=row+4, end_column=col_W)
        bw = ws.cell(row=row, column=col_W, value=combine_bikou_contract(rec))
        bw.alignment = center_vert; bw.font = red_font

        # 罫線・合計行の網掛け
        for c in range(2, col_W + 1):
            for r in range(row, row+5):
                ws.cell(row=r, column=c).border = thin_border
        for c in range(2, col_W + 1):
            ws.cell(row=row+4, column=c).fill = gray_fill

        blocks.append((row, row+4))
        row += 5

    # ---- 最下段「合計」行群（SUMIF） ----
    # 見本では「合　　　計」(B列) + 科目縦に4行（家賃/共益費/駐車料/水道料）
    sum_start = row
    ws.cell(row=sum_start, column=2, value="合　　　計")
    ws.cell(row=sum_start, column=4, value="家賃")
    for i, name in enumerate(["共益費　","駐車料","水道料"], start=1):
        ws.cell(row=sum_start + i, column=4, value=name)

    # D列の科目名をキーに SUMIF で E..T 列を集計
    # データ範囲は D{first_data_row}:D{last_data_row} で、first_data_row=6, last_data_row=row-1
    first_data_row = 6
    last_data_row  = row - 1
    def sumif_range(col_letter):
        return f"{col_letter}${first_data_row}:{col_letter}${last_data_row}"

    for i in range(4):
        r = sum_start + i
        # E..R （月列含む）+ S（横計）+ T（期末）
        for cidx in range(5, 7+num_months+1+1):  # E(5) .. T
            col_letter = get_column_letter(cidx)
            # 科目の見本セル (D列) を参照（例：$D$7 の文字列と一致するものを合計）
            # ここでは「その行の D セルと同じ科目」を集計
            ws.cell(row=r, column=cidx).number_format = number_fmt
            ws.cell(row=r, column=cidx).value = f"=SUMIF($D${first_data_row}:$D${last_data_row},$D${r},{sumif_range(col_letter)})"

    # U列（礼金・更新料）、V列（敷金）は単純合計
    col_U = 7 + num_months + 2
    col_V = col_U + 1
    for cidx in [col_U, col_V]:
        col_letter = get_column_letter(cidx)
        ws.cell(row=sum_start, column=cidx).number_format = number_fmt
        ws.cell(row=sum_start, column=cidx).value = f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})"
        # 下3行（共益費/駐車料/水道料）は空欄にしておく
        for i in range(1,4):
            ws.cell(row=sum_start+i, column=cidx, value=None)

    # 右端 備考は空欄
    for i in range(4):
        ws.cell(row=sum_start+i, column=col_V+1, value="")

    # 体裁
    for c in range(2, col_V+1):
        for r in range(sum_start, sum_start+4):
            ws.cell(row=r, column=c).border = thin_border

    # 備考列の幅（最大長に合わせる）
    ws.column_dimensions[get_column_letter(col_V+1)].width = max(
        [len(combine_bikou_contract(rec)) for rec in records] + [10]
    ) * 1.6

    # バイト列に保存
    out_file = io.BytesIO()
    wb.save(out_file)
    return out_file.getvalue()

# ========== Streamlit UI ==========
st.set_page_config(page_title="入居管理表アプリ", layout="wide")
st.title("📊 収支報告書PDFから入居管理表を作成（改修版）")

PASSWORD = st.secrets["APP_PASSWORD"]
pw = st.text_input("パスワードを入力してください", type="password")
if pw != PASSWORD:
    st.warning("正しいパスワードを入力してください。")
    st.stop()

property_name = st.text_input("物件名（例：XOヒルズ）", value="")
uploaded_files = st.file_uploader("収支報告書PDFを最大12ファイルまでアップロードしてください", type="pdf", accept_multiple_files=True)

if uploaded_files and st.button("入居管理表を作成"):
    if len(uploaded_files) > 12:
        st.warning("アップロードできるのは最大12ファイルまでです。")
    else:
        st.info("収支報告書を読み取り中...")
        records, months = asyncio.run(process_files(uploaded_files))
        if not records:
            st.error("データが抽出できませんでした。PDFの品質やフォーマットをご確認ください。")
            st.stop()

        st.info("入居管理表を作成中...")
        excel_data = export_excel(records, months, property_name)
        if months:
            start_month = months[0].replace("-", "年") + "月"
            end_month   = months[-1].replace("-", "年") + "月"
            fn = f"{property_name or '入居管理表'}（{start_month}〜{end_month}）_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
        else:
            fn = f"{property_name or '入居管理表'}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"

        st.download_button("入居管理表をダウンロード", data=excel_data,
                           file_name=fn,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success("完了しました。")

