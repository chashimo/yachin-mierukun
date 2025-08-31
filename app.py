# 入居管理表アプリ（v2 改修版）
# - 収入明細の「部屋行」だけを text_context に渡す抽出器を追加
# - Vision プロンプトを強化（列順固定／空欄は0で埋める／記号保持／収入以外を無視）
# - Excel 出力をサンプル体裁に合わせて更新（B6開始、物件名位置、総合計・確認用行、フリーズペイン 等）

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
    "あなたは不動産管理のOCR整形アシスタントです。"
    "入力PDFは不揃いで、0円の欄は印字されず列が“詰まって”見える場合があります。"
    "次の厳密なルールで『収入明細（部屋行）』を JSON 化してください。\n"
    "\n"
    "【必須ルール】\n"
    "1) 科目の列順は必ず固定：賃料(rent) → 共益費(fee) → 駐車料(parking) → 水道料(water)。\n"
    "2) どれかの値が空欄でも 0 を入れて 4科目を常に埋める（列を詰めない・飛ばさない）。\n"
    "3) 数値はカンマ無し整数。月キーは YYYY-MM。\n"
    "4) テナント名・記号（×・※・㈱・△ など）は削除・置換しない。\n"
    "5) 『収入明細』以外（支出明細・注記）は無視。曖昧な語に引っ張られない。\n"
    "\n"
    "【科目ラベリングの優先ヒント】\n"
    "A) 共益費(fee)は月をまたいで一定（例：2000, 5000, 11000 など）が多い。行内で家賃の右に現れる小さな定額は fee を優先。\n"
    "B) 駐車料(parking)は Pxx 行（P01/P02…）や『（0001）込駐車場』等の備考と論理的に結びつく。根拠が弱い場合は fee を優先。\n"
    "C) 水道料(water)は典型的には 0。明確な『水道』文脈があるときのみ water とする。\n"
    "\n"
    "【出力構造】\n"
    "{\n"
    "  \"records\": [\n"
    "    {\n"
    "      \"room\": \"0102\" / \"P01\",\n"
    "      \"tenant\": \"氏名や法人名（記号含む）\",\n"
    "      \"monthly\": {\n"
    "        \"YYYY-MM\": {\"rent\":0, \"fee\":0, \"parking\":0, \"water\":0, \"reikin\":0, \"koushin\":0, \"bikou\":\"\"}\n"
    "      },\n"
    "      \"shikikin\": 0,\n"
    "      \"linked_room\": \"0001\"  # Pxx→住戸の手掛かりがあれば\n"
    "    }\n"
    "  ]\n"
    "}\n"
    "JSON 以外の文字（前置き・コードブロック）を出力しないでください。"
)

async def call_openai_vision_async(base64_images, text_context, default_month_id):
    image_parts = [{"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}} for b64 in base64_images]
    messages = [
        {"role": "system", "content": VISION_INSTRUCTIONS},
        {"role": "user", "content": [
            *image_parts,
            {"type": "text", "text":
                "【収入明細の部屋行のみのテキスト（0円は印字されず列が詰まる場合があります）】\n"
                + text_context +
                "\n\n出力は部屋ごと・YYYY-MMごとして、上記の列順を必ず守り、空欄は0で埋めてください。"
            }
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

# ---- 収入明細の「部屋行」だけを抽出 ----
INCOME_START_RE = re.compile(r"(収入の部|収入明細)")
EXPENSE_START_RE = re.compile(r"(支出の部|支出明細)")
ROOM_LINE_RE    = re.compile(r"^\s*(\d{3,4}|P\d{1,2})\s")   # 0001, 0102, 0303, P01 など

def extract_income_text_only(pdf_bytes: bytes) -> str:
    """PDFテキストから『収入明細の部屋行』だけを抽出し連結する。"""
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            raw = page.extract_text() or ""
            # HIDDEN/YNS_* 等のノイズ除去
            raw = re.sub(r"\b(?:YNS_[A-Z0-9_]+|HIDDEN[_A-Z0-9]+)\b", "", raw)

            # 収入明細〜支出明細の範囲に絞る
            sub = raw
            m1 = INCOME_START_RE.search(raw)
            m2 = EXPENSE_START_RE.search(raw)
            if m1:
                sub = raw[m1.end(): m2.start() if m2 else len(raw)]

            # 行頭が「部屋番号 or Pxx」で始まる行のみ採用
            lines = []
            for line in sub.splitlines():
                if ROOM_LINE_RE.match(line.strip()):
                    lines.append(line.strip())
            if lines:
                parts.append("\n".join(lines))
    return "\n".join(parts)

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

    # --- 収入明細の「部屋行」だけに絞った text_context ---
    text_context = extract_income_text_only(file_bytes)

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
      record = { room, tenant, monthly: { 'YYYY-MM': {...} }, shikikin, linked_room }
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
        # 敷金は最大
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
                if dst["bikou"]:
                    if b not in dst["bikou"]:
                        dst["bikou"] += f", {b}"
                else:
                    dst["bikou"] = b

def fold_parking_Pxx(all_recs):
    """Pxx レコードを、linked_room に駐車料として付替える（備考は契約単位で一意化）。"""
    to_delete = []
    # 検索用: room -> keys(list)
    by_room = {}
    for key, rec in all_recs.items():
        by_room.setdefault(rec["room"], []).append(key)

    p_sources_by_target = {}  # target_room -> set(["P01","P02",...])

    for key, rec in list(all_recs.items()):
        room = rec["room"]
        if not room.upper().startswith("P"):
            continue
        target_room = rec.get("linked_room") or ""
        if not target_room:
            # 備考から (dddd) を拾う fallback
            for mk, mv in rec.get("monthly", {}).items():
                m = re.search(r"（?(\d{3,4})）?", mv.get("bikou",""))
                if m:
                    target_room = m.group(1).zfill(4)
                    break
        if not target_room:
            logger.info(f"Pxx行 {key} は付替え先不明のため残存")
            continue

        target_keys = by_room.get(target_room, [])
        if not target_keys:
            tkey = (target_room, "")
            all_recs[tkey] = {"room": target_room, "tenant":"", "monthly": {}, "shikikin":0, "linked_room":""}
            by_room.setdefault(target_room, []).append(tkey)
            target_keys = [tkey]

        tkey = target_keys[0]
        target = all_recs[tkey]
        for mk, mv in rec.get("monthly", {}).items():
            dst = target["monthly"].setdefault(mk, {"rent":0,"fee":0,"parking":0,"water":0,"reikin":0,"koushin":0,"bikou":""})
            dst["parking"] += clean_int(mv.get("parking"))
        p_sources_by_target.setdefault(target_room, set()).add(room.upper())
        to_delete.append(key)

    for key in to_delete:
        all_recs.pop(key, None)

    # 備考を契約単位で一意に付与
    for (room, tenant), rec in all_recs.items():
        srcs = p_sources_by_target.get(room)
        if srcs:
            note = f"駐車場({','.join(sorted(srcs))})→{room}"
            for mv in rec.get("monthly", {}).values():
                b = mv.get("bikou") or ""
                if note not in b:
                    mv["bikou"] = (b + ", " if b else "") + note

async def process_files(files):
    tasks = [handle_file(file) for file in files]
    results = await asyncio.gather(*tasks)

    # 1) マージ
    all_recs = {}  # key = (room, tenant)
    for recs in results:
        merge_records(all_recs, recs)

    # 2) Pxx 付替え
    fold_parking_Pxx(all_recs)

    # 3) 出力用整形
    out_sorted = sorted(
        all_recs.values(),
        key=lambda r: (
            (9000 + int(re.sub(r"\D","",r["room"]) or 0)) if r["room"].upper().startswith("P")
            else int(re.sub(r"\D","",r["room"]) or 0),
            r["tenant"] or "~",
            sorted(r["monthly"].keys())[0] if r["monthly"] else "9999-99"
        )
    )

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

    header_row = 6           # ヘッダ行（=表の左上は B6）
    data_start_row = 7       # データ開始行
    number_fmt  = "#,##0"

    header_fill = PatternFill("solid", fgColor="BDD7EE")
    green_fill  = PatternFill("solid", fgColor="CCFFCC")
    yellow_fill = PatternFill("solid", fgColor="FFF2CC")  # 合計行
    pink_fill   = PatternFill("solid", fgColor="F8CBAD")  # 総合計行
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    center_vert = Alignment(vertical="center", wrap_text=True)
    bold_font   = Font(bold=True)
    red_font    = Font(color="9C0000")
    thin_border = Border(*[Side(style='thin')] * 4)
    thick       = Side(style="thick")
    thick_border = Border(left=thick, right=thick, top=thick, bottom=thick)

    num_months = len(months)
    col_B = 2; col_C = 3; col_D = 4; col_E = 5; col_F = 6; col_G = 7
    col_month_end = 6 + num_months
    col_S = col_month_end + 1
    col_T = col_month_end + 2
    col_U = col_month_end + 3
    col_V = col_month_end + 4
    col_W = col_month_end + 5
    col_X = col_W + 1

    # ---- タイトル（物件名は含めない） ----
    ws.merge_cells(start_row=2, start_column=col_B, end_row=2, end_column=col_W)
    if months:
        start_month = months[0].replace("-", "年") + "月"
        end_month   = months[-1].replace("-", "年") + "月"
        ws.cell(row=2, column=col_B, value=f"入居管理表 （{start_month}〜{end_month}）")
    else:
        ws.cell(row=2, column=col_B, value="入居管理表")
    ws.cell(row=2, column=col_B).font = Font(size=14, bold=True)
    ws.cell(row=2, column=col_B).alignment = center

    # 物件名ラベル＆値
    ws.merge_cells(start_row=4, start_column=col_B, end_row=4, end_column=col_C)
    ws.cell(row=4, column=col_B, value="物件名").alignment = center
    ws.merge_cells(start_row=4, start_column=col_D, end_row=4, end_column=col_F)
    ws.cell(row=4, column=col_D, value=(property_name or "")).alignment = center
    for addr in ("B4","C4","D4","E4","F4"):
        ws[addr].border = thick_border

    # ---- ヘッダ（B6..）----
    ws.merge_cells(start_row=header_row, start_column=col_B, end_row=header_row, end_column=col_C)
    ws.cell(row=header_row, column=col_B, value="賃借人")
    ws.merge_cells(start_row=header_row, start_column=col_D, end_row=header_row, end_column=col_E)
    ws.cell(row=header_row, column=col_D, value="基準額")
    ws.cell(row=header_row, column=col_F, value="期首\n未収/前受")
    for i, m in enumerate(months):
        mm = int(m[5:])
        ws.cell(row=header_row, column=col_G+i, value=f"{mm}月")
    ws.cell(row=header_row, column=col_S, value="合計")
    ws.cell(row=header_row, column=col_T, value="期末\n未収/前受")
    ws.cell(row=header_row, column=col_U, value="礼金・更新料")
    ws.cell(row=header_row, column=col_V, value="敷金")
    ws.cell(row=header_row, column=col_W, value="備考")
    for c in range(col_B, col_W+1):
        cc = ws.cell(row=header_row, column=c)
        cc.fill = header_fill; cc.font = bold_font; cc.alignment = center

    # ---- データ（5行ブロック）----
    row = data_start_row
    for rec in records:
        room   = rec.get("room","")
        tenant = rec.get("tenant","")
        # 左側（室番号/賃借人）
        ws.merge_cells(start_row=row,   start_column=col_B, end_row=row+4, end_column=col_B)
        ws.cell(row=row, column=col_B, value="室番号").alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        ws.cell(row=row, column=col_C, value=room).alignment = center
        ws.cell(row=row, column=col_C).fill = green_fill
        ws.merge_cells(start_row=row+1, start_column=col_C, end_row=row+4, end_column=col_C)
        ws.cell(row=row+1, column=col_C, value=tenant).alignment = center

        # 科目（D列）と基準額（E列）。基準額は現状 0 初期化（必要に応じて設定）
        subjects = ["家賃","共益費　","駐車料","水道料","合計"]
        for i, s in enumerate(subjects):
            ws.cell(row=row+i, column=col_D, value=s)
        for i, v in enumerate([0, 0, 0, 0]):
            cc = ws.cell(row=row+i, column=col_E, value=v); cc.number_format = number_fmt
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
            ws.cell(row=row+4, column=col_G+i, value=f"=SUM({get_column_letter(col_G+i)}{row}:{get_column_letter(col_G+i)}{row+3})").number_format = number_fmt

        # 横計 S列
        for r_i in range(5):
            ws.cell(row=row+r_i, column=col_S, value=f"=SUM({get_column_letter(col_G)}{row+r_i}:{get_column_letter(col_month_end)}{row+r_i})").number_format = number_fmt

        # 期末 T列
        for r_i in range(4):
            ws.cell(row=row+r_i, column=col_T, value=0).number_format = number_fmt
        ws.cell(row=row+4, column=col_T, value=f"=SUM({get_column_letter(col_T)}{row}:{get_column_letter(col_T)}{row+3})").number_format = number_fmt

        # U: 礼金・更新料（5行結合） / V: 敷金（5行結合） / W: 備考（5行結合）
        uval = sum((mv.get("reikin",0)+mv.get("koushin",0)) for mv in rec.get("monthly",{}).values())
        ws.merge_cells(start_row=row, start_column=col_U, end_row=row+4, end_column=col_U)
        ws.cell(row=row, column=col_U, value=uval).alignment = center_vert
        ws.cell(row=row, column=col_U).number_format = number_fmt

        ws.merge_cells(start_row=row, start_column=col_V, end_row=row+4, end_column=col_V)
        ws.cell(row=row, column=col_V, value=rec.get("shikikin",0)).alignment = center_vert
        ws.cell(row=row, column=col_V).number_format = number_fmt

        ws.merge_cells(start_row=row, start_column=col_W, end_row=row+4, end_column=col_W)
        bn = ws.cell(row=row, column=col_W, value=combine_bikou_contract(rec))
        bn.alignment = center_vert; bn.font = red_font

        # 罫線・合計行の色
        for c in range(col_B, col_W+1):
            for r in range(row, row+5):
                ws.cell(row=r, column=c).border = thin_border
        for c in range(col_B, col_W+1):
            ws.cell(row=row+4, column=c).fill = yellow_fill

        row += 5

    first_data_row = data_start_row
    last_data_row  = row - 1

    # ---- 下段「合計」4行（家賃/共益費/駐車料/水道料） ----
    sum_start = row
    ws.merge_cells(start_row=sum_start, end_row=sum_start+3, start_column=col_B, end_column=col_C)
    ws.cell(row=sum_start, column=col_B, value="合計").alignment = center

    for i, name in enumerate(["家賃","共益費　","駐車料","水道料"]):
        ws.cell(row=sum_start+i, column=col_D, value=name)

    def sumif_range(col_letter):
        return f"{col_letter}${first_data_row}:{col_letter}${last_data_row}"

    for i in range(4):
        r = sum_start + i
        for cidx in range(col_E, col_T+1):  # E..T
            col_letter = get_column_letter(cidx)
            ws.cell(row=r, column=cidx, value=f"=SUMIF($D${first_data_row}:$D${last_data_row},$D${r},{sumif_range(col_letter)})").number_format = number_fmt
        # 罫線
        for c in range(col_B, col_W+1):
            ws.cell(row=r, column=c).border = thin_border

    # U/V は全データの単純合計（最上段のみ）
    for cidx in [col_U, col_V]:
        col_letter = get_column_letter(cidx)
        ws.cell(row=sum_start, column=cidx, value=f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})").number_format = number_fmt
        for i in range(1,4):
            ws.cell(row=sum_start+i, column=cidx, value=None)

    # ---- 最終行「総合計」 ----
    grand_row = sum_start + 4
    ws.merge_cells(start_row=grand_row, end_row=grand_row, start_column=col_B, end_column=col_C)
    ws.cell(row=grand_row, column=col_B, value="総合計").alignment = center
    for cidx in range(col_E, col_T+1):
        col_letter = get_column_letter(cidx)
        ws.cell(row=grand_row, column=cidx, value=f"=SUM({col_letter}{sum_start}:{col_letter}{sum_start+3})").number_format = number_fmt
    for cidx in [col_U, col_V]:
        col_letter = get_column_letter(cidx)
        ws.cell(row=grand_row, column=cidx, value=f"=SUM({col_letter}{sum_start}:{col_letter}{sum_start})").number_format = number_fmt
    for c in range(col_B, col_W+1):
        ws.cell(row=grand_row, column=c).fill = pink_fill
        ws.cell(row=grand_row, column=c).border = thin_border

    # ---- 右外側「確認用」 & 一括チェック式 ----
    ws.cell(row=grand_row-1, column=col_X, value="確認用").alignment = center
    g_letter = get_column_letter(col_G)
    r_letter = get_column_letter(col_month_end)
    ws.cell(row=grand_row, column=col_X, value=f"=SUM({g_letter}{first_data_row}:{r_letter}{last_data_row})/2").number_format = number_fmt

    # ---- 2行下の「算式確認」行 ----
    check_row = grand_row + 2
    ws.cell(row=check_row, column=col_E, value="算式確認")
    for cidx in range(col_F, col_T+1):  # F..T
        col_letter = get_column_letter(cidx)
        ws.cell(row=check_row, column=cidx, value=f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})/2").number_format = number_fmt

    # 備考列の幅（可変）
    ws.column_dimensions[get_column_letter(col_W)].width = max(
        [len(combine_bikou_contract(rec)) for rec in records] + [10]
    ) * 1.6

    # フリーズペイン（C/D と 6/7 の境界）
    try:
        ws.freeze_panes = ws.cell(row=data_start_row, column=4)  # "D7"
    except Exception:
        pass

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ========== Streamlit UI ==========
st.set_page_config(page_title="入居管理表アプリ", layout="wide")
st.title("📊 収支報告書PDFから入居管理表を作成（v2改修版）")

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

