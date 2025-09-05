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
    ws.title = property_name or "入居管理表"

    # ---- 定数・スタイル ----
    header_row = 6           # ヘッダ行（=表の左上は B6）
    data_start_row = 7       # データ開始行
    last_fixed_col = 3       # C列まで固定 → freeze_panes="D7" で列＆行を同時固定
    number_fmt  = "#,##0"

    header_fill = PatternFill("solid", fgColor="BDD7EE")
    green_fill  = PatternFill("solid", fgColor="CCFFCC")
    gray_fill   = PatternFill("solid", fgColor="DDDDDD")
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    center_vert = Alignment(vertical="center", wrap_text=True)
    bold_font   = Font(bold=True)
    red_font    = Font(color="9C0000")
    thin_border = Border(*[Side(style='thin')] * 4)

	# 合計=黄、総合計=ピンク、太線枠
    yellow_fill = PatternFill("solid", fgColor="FFF2CC")   # 合計
    pink_fill   = PatternFill("solid", fgColor="F8CBAD")   # 総合計
    thick_side  = Side(style="thick")
    thick_border = Border(left=thick_side, right=thick_side, top=thick_side, bottom=thick_side)
    
    num_months = len(months)
    # 列インデックス
    col_B = 2
    col_C = 3
    col_D = 4
    col_E = 5
    col_F = 6
    col_G = 7
    col_month_end = 6 + num_months        # G..(6+num_months)
    col_S = col_month_end + 1             # 合計
    col_T = col_month_end + 2             # 期末 未収/前受
    col_U = col_month_end + 3             # 礼金・更新料
    col_V = col_month_end + 4             # 敷金
    col_W = col_month_end + 5             # 備考
    col_X = col_W + 1                     # 備考の一つ右（確認用の欄）

    # ---- タイトル & 物件名 ----
    # B2：タイトル（物件名は入れない）
    ws.merge_cells(start_row=2, start_column=col_B, end_row=2, end_column=col_W)
    if months:
        start_month = months[0].replace("-", "年") + "月"
        end_month   = months[-1].replace("-", "年") + "月"
        ws.cell(row=2, column=col_B, value=f"入居管理表 （{start_month}〜{end_month}）")
    else:
        ws.cell(row=2, column=col_B, value="入居管理表")
    ws.cell(row=2, column=col_B).font = Font(size=14, bold=True)
    ws.cell(row=2, column=col_B).alignment = center

    # B4:C4 = 物件名, D4:F4 = 物件名の値
    ws.merge_cells(start_row=4, start_column=col_B, end_row=4, end_column=col_C)
    ws.cell(row=4, column=col_B, value="物件名").alignment = center
    ws.merge_cells(start_row=4, start_column=col_D, end_row=4, end_column=col_F)
    ws.cell(row=4, column=col_D, value=(property_name or "")).alignment = center

    # 太線罫線
    for c in range(col_B, col_C+1):  # B4:C4
        ws.cell(row=4, column=c).border = thick_border
    for c in range(col_D, col_F+1):  # D4:F4
        ws.cell(row=4, column=c).border = thick_border

    # ---- ヘッダ（B6..）----
    ws.merge_cells(start_row=header_row, start_column=col_B, end_row=header_row, end_column=col_C)
    ws.cell(row=header_row, column=col_B, value="賃借人")

    ws.merge_cells(start_row=header_row, start_column=col_D, end_row=header_row, end_column=col_E)
    ws.cell(row=header_row, column=col_D, value="基準額")

    ws.cell(row=header_row, column=col_F, value="期首\n未収/前受")

    # 月見出し G..（数は動的）
    for i, m in enumerate(months):
        mm = int(m[5:])
        ws.cell(row=header_row, column=col_G+i, value=f"{mm}月")

    ws.cell(row=header_row, column=col_S, value="合計")
    ws.cell(row=header_row, column=col_T, value="期末\n未収/前受")
    ws.cell(row=header_row, column=col_U, value="礼金・更新料")
    ws.cell(row=header_row, column=col_V, value="敷金")
    ws.cell(row=header_row, column=col_W, value="備考")

    # ヘッダの体裁
    for c in range(col_B, col_W+1):
        cc = ws.cell(row=header_row, column=c)
        cc.fill = header_fill
        cc.font = bold_font
        cc.alignment = center

    # ---- データ（5行ブロック）----
    row = data_start_row
    blocks = []  # (start_row, end_row) for each 5-row block
    for rec in records:
        room   = rec.get("room","")
        tenant = rec.get("tenant","")
        base_r = rec.get("base_rent",0)
        base_f = rec.get("base_fee",0)
        base_p = rec.get("base_parking",0)
        base_w = rec.get("base_water",0)
        shikikin = rec.get("shikikin",0)
        reikin_koushin_total = sum((mv.get("reikin",0)+mv.get("koushin",0)) for mv in rec.get("monthly",{}).values())

        # 左側（室番号/賃借人）
        ws.merge_cells(start_row=row,   start_column=col_B, end_row=row+4, end_column=col_B)
        ws.cell(row=row, column=col_B, value="室番号").alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        ws.cell(row=row, column=col_C, value=room).alignment = center
        ws.cell(row=row, column=col_C).fill = green_fill
        ws.merge_cells(start_row=row+1, start_column=col_C, end_row=row+4, end_column=col_C)
        ws.cell(row=row+1, column=col_C, value=tenant).alignment = center

        # 科目（D列）と基準額（E列）
        subjects = ["家賃","共益費　","駐車料","水道料","合計"]
        for i, s in enumerate(subjects):
            ws.cell(row=row+i, column=col_D, value=s)
        for i, v in enumerate([base_r, base_f, base_p, base_w]):
            cc = ws.cell(row=row+i, column=col_E, value=v); cc.number_format = number_fmt
        # E列 合計は式
        ws.cell(row=row+4, column=col_E, value=f"=SUM(E{row}:E{row+3})").number_format = number_fmt

        # 期首（F列）
        for i in range(5):
            ws.cell(row=row+i, column=col_F, value=0).number_format = number_fmt
        ws.cell(row=row+4, column=col_F, value=f"=SUM(F{row}:F{row+3})").number_format = number_fmt

        # 月次 G..
        for i, m in enumerate(months):
            mv = (rec.get("monthly") or {}).get(m, {})
            vals = [mv.get("rent",0), mv.get("fee",0), mv.get("parking",0), mv.get("water",0)]
            for r_i, v in enumerate(vals):
                cc = ws.cell(row=row+r_i, column=col_G+i, value=v)
                cc.number_format = number_fmt
            # 月次の「合計」行（5行目）は縦計式
            ws.cell(row=row+4, column=col_G+i, value=f"=SUM({get_column_letter(col_G+i)}{row}:{get_column_letter(col_G+i)}{row+3})").number_format = number_fmt

        # 横計 S列
        for r_i in range(5):
            ws.cell(row=row+r_i, column=col_S, value=f"=SUM({get_column_letter(col_G)}{row+r_i}:{get_column_letter(col_month_end)}{row+r_i})").number_format = number_fmt

        # 期末 T列
        for r_i in range(4):
            ws.cell(row=row+r_i, column=col_T, value=0).number_format = number_fmt
        ws.cell(row=row+4, column=col_T, value=f"=SUM({get_column_letter(col_T)}{row}:{get_column_letter(col_T)}{row+3})").number_format = number_fmt

        # U: 礼金・更新料（5行結合）
        ws.merge_cells(start_row=row, start_column=col_U, end_row=row+4, end_column=col_U)
        cu = ws.cell(row=row, column=col_U, value=reikin_koushin_total); cu.alignment = center_vert; cu.number_format = number_fmt
        # V: 敷金（5行結合）
        ws.merge_cells(start_row=row, start_column=col_V, end_row=row+4, end_column=col_V)
        cv = ws.cell(row=row, column=col_V, value=shikikin); cv.alignment = center_vert; cv.number_format = number_fmt
        # W: 備考（5行結合）
        ws.merge_cells(start_row=row, start_column=col_W, end_row=row+4, end_column=col_W)
        bw = ws.cell(row=row, column=col_W, value=combine_bikou_contract(rec)); bw.alignment = center_vert; bw.font = red_font

        # 罫線・黄色網掛け（ブロック内）
        for c in range(col_B, col_W+1):
            for r in range(row, row+5):
                ws.cell(row=r, column=c).border = thin_border
        for c in range(col_B, col_W+1):
            ws.cell(row=row+4, column=c).fill = yellow_fill

        blocks.append((row, row+4))
        row += 5

    # データ範囲（合計などの式用）
    first_data_row = data_start_row
    last_data_row  = row - 1  # データの最終行（ブロック終端）

    # ---- 下段「合計」4行（家賃/共益費/駐車料/水道料） ----
    sum_start = row
    # B..C を4行縦結合して「合計」
    ws.merge_cells(start_row=sum_start, end_row=sum_start+3, start_column=col_B, end_column=col_C)
    ws.cell(row=sum_start, column=col_B, value="合計").alignment = center

    # 科目名（D列）
    for i, name in enumerate(["家賃","共益費　","駐車料","水道料"]):
        ws.cell(row=sum_start+i, column=col_D, value=name)

    # D列の科目名をキーに、E..T を SUMIF で縦集計
    def sumif_range(col_letter):
        return f"{col_letter}${first_data_row}:{col_letter}${last_data_row}"
    for i in range(4):
        r = sum_start + i
        for cidx in range(col_E, col_T+1):  # E..T
            col_letter = get_column_letter(cidx)
            ws.cell(row=r, column=cidx, value=f"=SUMIF($D${first_data_row}:$D${last_data_row},$D${r},{sumif_range(col_letter)})").number_format = number_fmt

    # U/V は全データの単純合計（最上段のみ表示、下2〜4行は空欄）
    for cidx in [col_U, col_V]:
        col_letter = get_column_letter(cidx)
        ws.cell(row=sum_start, column=cidx, value=f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})").number_format = number_fmt
        for i in range(1,4):
            ws.cell(row=sum_start+i, column=cidx, value=None)

    # 備考列は空欄
    for i in range(4):
        ws.cell(row=sum_start+i, column=col_W, value="")

    # 体裁
    for c in range(col_B, col_W+1):
        for r in range(sum_start, sum_start+4):
            ws.cell(row=r, column=c).border = thin_border

    # ---- 最終行「総合計」 ----
    grand_row = sum_start + 4
    # 見出し（B..Cは横1行なので結合は任意。合わせて結合しておく）
    ws.merge_cells(start_row=grand_row, end_row=grand_row, start_column=col_B, end_column=col_C)
    ws.cell(row=grand_row, column=col_B, value="総合計").alignment = center
    # E..T は上の4行合算（=SUM(同列の合計4行分)）
    for cidx in range(col_E, col_T+1):
        col_letter = get_column_letter(cidx)
        ws.cell(row=grand_row, column=cidx, value=f"=SUM({col_letter}{sum_start}:{col_letter}{sum_start+3})").number_format = number_fmt
    # U/V も合算
    for cidx in [col_U, col_V]:
        col_letter = get_column_letter(cidx)
        ws.cell(row=grand_row, column=cidx, value=f"=SUM({col_letter}{sum_start}:{col_letter}{sum_start})").number_format = number_fmt  # 上段のみ値が入る

    # 総合計罫線、ピンク
    for c in range(col_B, col_W+1):
        ws.cell(row=grand_row, column=c).border = thin_border
        ws.cell(row=grand_row, column=c).fill = pink_fill

    # ---- 右外側「確認用」 & 一括チェック式（8）----
    ws.cell(row=grand_row-1, column=col_X, value="確認用").alignment = center
    g_letter = get_column_letter(col_G)
    r_letter = get_column_letter(col_month_end)
    ws.cell(row=grand_row, column=col_X, value=f"=SUM({g_letter}{first_data_row}:{r_letter}{last_data_row})/2").number_format = number_fmt

    # ---- 2行下の「算式確認」行（9）----
    check_row = grand_row + 2
    ws.cell(row=check_row, column=col_E, value="算式確認")
    for cidx in range(col_F, col_T+1):  # F..T
        col_letter = get_column_letter(cidx)
        ws.cell(row=check_row, column=cidx, value=f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})/2").number_format = number_fmt

    # 備考列の幅（可変）
    ws.column_dimensions[get_column_letter(col_W)].width = max(
        [len(combine_bikou_contract(rec)) for rec in records] + [10]
    ) * 1.6

    # ---- ウィンドウ枠の固定（4,5）----
    try:
        ws.freeze_panes = ws.cell(row=data_start_row, column=last_fixed_col+1)  # "D7" 相当
        # → 左に C まで・上に 6 行目まで固定
    except Exception:
        pass  # 固定できなくても実害が出ないように

    # 保存
    import io
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


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

