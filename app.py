# 入居管理表アプリ（pdfplumber-only 完全版）
# - 2ページ目の「収入明細」テーブルを pdfplumber で抽出
# - Vision フォールバックなし（失敗したらエラー）
# - 「一番上のテーブル＝収入」を採用
# - 非空率しきい値 0.01、ヘッダは最初の行固定、ゼロ行＆備考空もスキップしない
# - Excel禁止文字（\x00 など）を xls_clean() で除去し IllegalCharacterError を回避
# - fold_parking_Pxx は契約者名で居室へ付け替え
# - 基準額は最頻値（モード）で決定

import streamlit as st
import io
import json
import asyncio
import re
import logging
from datetime import datetime
from pathlib import Path
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
import subprocess
from collections import Counter

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

# ===== サイドバー：ブランチ/コミット表示（混線防止） =====
try:
    BRANCH = subprocess.check_output(["git","rev-parse","--abbrev-ref","HEAD"], text=True).strip()
    COMMIT = subprocess.check_output(["git","rev-parse","--short","HEAD"], text=True).strip()
    st.sidebar.info(f"branch: {BRANCH}\ncommit: {COMMIT}")
except Exception as e:
    st.sidebar.warning(f"git情報を取得できませんでした: {e}")

# ===== Excel禁止文字の除去ユーティリティ =====
def xls_clean(v):
    if v is None:
        return None
    s = str(v)
    s = ILLEGAL_CHARACTERS_RE.sub("", s)  # openpyxl の禁止文字（制御文字など）を除去
    return s

# ========== 文字列/数値ユーティリティ ==========
def _normalize_cell(s):
    if s is None: return ""
    s = str(s).replace("\x00","").strip()   # NULL 早期除去
    return s

def _to_int_like(s):
    s = _normalize_cell(s).replace(",", "").replace("¥", "").replace("円", "")
    if s == "": return 0
    try: return int(float(s))
    except: return 0

def extract_text_with_pdfplumber(pdf_bytes: bytes) -> str:
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

# 備考をカンマ区切りで保持しつつ、重複を除去して結合
def append_note_unique(current: str, note: str) -> str:
    """
    current: 既存の備考（'a, b' など）
    note   : 追加したい備考（'b' など）
    返り値: 重複を取り除き、順序維持で 'a, b, c' の形にして返す
    """
    if not note:
        return current or ""
    # 区切りの統一（全角読点も受ける）、前後スペース除去
    def _tok(s: str):
        return [t for t in re.split(r"[,\u3001]\s*", (s or "")) if t]

    tokens = _tok(current)
    seen = set(tokens)
    n = note.strip()
    if n and n not in seen:
        tokens.append(n)
    return ", ".join(tokens)

# ====== pdfplumber 抽出（収入明細テーブル：最上段のみ採用） ======
def extract_income_table_with_pdfplumber(pdf_bytes: bytes, top_margin_px: int = 40, side_margin_px: int = 24):
    """
    2ページ目のROIを切って、最初（最上段）のテーブル（=収入明細）を返す。
    返り値: 2Dリスト（行×列） / None（失敗）
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if len(pdf.pages) == 0:
                return None
            page = pdf.pages[1] if len(pdf.pages) >= 2 else pdf.pages[0]

            # テキストPDFか簡易判定（スキャンは失敗扱い）
            if not page.chars or len(page.chars) < 10:
                return None

            W, H = page.width, page.height
            px_to_pt = 0.75  # ≒ 96dpi → pt
            x0 = side_margin_px * px_to_pt
            x1 = W - x0
            y0 = top_margin_px * px_to_pt
            y1 = H - (12 * px_to_pt)  # 下端に少し余白
            crop = page.crop((x0, y0, x1, y1))

            # ① 罫線ベース
            lattice = {
                "vertical_strategy":"lines","horizontal_strategy":"lines",
                "snap_tolerance":3,"join_tolerance":3,
                "intersection_x_tolerance":5,"intersection_y_tolerance":5,
                "edge_min_length":30,
            }
            tables = crop.extract_tables(lattice)

            # ② 文字整列ベース
            if not tables:
                stream = {
                    "vertical_strategy":"text","horizontal_strategy":"text",
                    "text_x_tolerance":2,"text_y_tolerance":2,
                    "snap_tolerance":3,"join_tolerance":3,
                }
                tables = crop.extract_tables(stream)

            if not tables:
                return None

            # 「一番上のテーブル＝収入」を採用
            table = tables[0]
            cleaned = [[_normalize_cell(c) for c in row] for row in table]

            # 非空率しきい値を 0.01 に緩和
            cells = sum(len(r) for r in cleaned)
            nonempty = sum(1 for r in cleaned for c in r if c)
            if cells == 0 or (nonempty / cells) < 0.01:
                return None
            return cleaned
    except Exception as e:
        logger.warning(f"pdfplumber抽出で例外: {e}")
        return None

def parse_income_table_to_records(table_2d, default_month_id: str):
    """
    ヘッダは table_2d の最初の行に固定。
    列名マッピングは提示ヘッダに準拠。
    ・部屋       → room
    ・契約者     → tenant
    ・年／月     → 月キー（YYYY-MM化を試み、失敗時は default_month_id）
    ・賃料/共益費/駐車料/水道代/礼金/更新料/備考 → monthly の対応項目へ
    ※ ゼロ行・備考空でもスキップしない（そのまま出力）
    """
    if not table_2d or len(table_2d) < 2:
        return []

    # 先頭行がヘッダ（空白除去のみ）
    headers = [re.sub(r"\s+", "", h) for h in table_2d[0]]
    data_rows = table_2d[1:]

    def find_col(*names):
        for n in names:
            if n in headers:
                return headers.index(n)
        return None

    COL = {
        "room":    find_col("部屋","室","号室","室番号"),
        "tenant":  find_col("契約者","賃借人","入居者","名義"),
        "month":   find_col("年／月","年/月","年月"),
        "rent":    find_col("賃料","家賃"),
        "fee":     find_col("共益費","管理費"),
        "parking": find_col("駐車料","Ｐ","Ｐ料金","P料金"),
        "water":   find_col("水道代","水道料"),
        "reikin":  find_col("礼金"),
        "koushin": find_col("更新料"),
        "bikou":   find_col("備考","摘要","特記事項"),
    }

    def at(row, key):
        j = COL.get(key)
        if j is None or j >= len(row): return ""
        return _normalize_cell(row[j])

    def month_from_cell(cell_value: str, fallback: str) -> str:
        """
        '2025/08' '2025-8' '2025年8月' '25/8' 等を YYYY-MM へ。
        失敗時は fallback（月キー正規化は month_key() に委ねる）
        """
        s = str(cell_value or "").strip()
        if not s:
            return month_key(fallback)
        m = re.search(r'(\d{2,4})[年/\-\.](\d{1,2})', s)
        if m:
            yy = m.group(1)
            mm = m.group(2).zfill(2)
            if len(yy) == 2:
                yy = "20" + yy
            return f"{yy}-{mm}"
        return month_key(fallback)

    records = []
    for row in data_rows:
        # ① 集計行（合計/総合計）を除外
        room_raw   = at(row, "room")
        tenant_raw = at(row, "tenant")
        # 「合　計」「総　計」のような全角スペース入りも拾う
        if re.search(r"合\s*計|総\s*計|合計額|総合計", room_raw) or \
                re.search(r"合\s*計|総\s*計|合計額|総合計", tenant_raw):
                    continue

        mk = month_from_cell(at(row, "month"), default_month_id)
        rec = {
            "room":        normalize_room(at(row, "room")),
            "tenant":      (at(row, "tenant") or "").strip(),
            "monthly": {
                mk: {
                    "rent":    clean_int(_to_int_like(at(row, "rent"))),
                    "fee":     clean_int(_to_int_like(at(row, "fee"))),
                    "parking": clean_int(_to_int_like(at(row, "parking"))),
                    "water":   clean_int(_to_int_like(at(row, "water"))),
                    "reikin":  clean_int(_to_int_like(at(row, "reikin"))),
                    "koushin": clean_int(_to_int_like(at(row, "koushin"))),
                    "bikou":   at(row, "bikou"),
                }
            },
            "shikikin":    0,
            "linked_room": "",
        }
        records.append(rec)

    return records

# ========== 1ファイル処理（pdfplumberオンリー） ==========
async def handle_file(file):
    file_name = file.name
    logger.info(f"開始: {file_name}")
    default_month_id = extract_month_from_filename(file_name)
    file_bytes = file.read()

    # pdfplumberのみ（失敗したら終了）
    table = extract_income_table_with_pdfplumber(file_bytes)
    if not table:
        st.error(f"{file_name}: 収入明細の表を検出できませんでした（pdfplumber）。")
        logger.error(f"{file_name}: pdfplumber抽出失敗（テキストPDFでない/表検出不可/非空率低）")
        return []

    try:
        records = parse_income_table_to_records(table, default_month_id)
        if not records:
            st.error(f"{file_name}: 表は見つかりましたが、明細のパースに失敗しました。")
            logger.error(f"{file_name}: テーブル→records 変換に失敗（列マッピング/数値化の不一致）")
            return []
        logger.info(f"{file_name}: pdfplumber抽出成功 / {len(records)}件")
        return records
    except Exception as e:
        st.error(f"{file_name}: 明細のパースで例外が発生しました。")
        logger.exception(e)
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
                dst["bikou"] = append_note_unique(dst.get("bikou"), b)

def fold_parking_Pxx(all_recs):
    """
    Pxx レコードを、契約者名（tenant）で一致する部屋のレコードへ駐車料として付替える。
    - tenant が空なら付替え不可で残す
    - 同名契約者の居室（room が P 以外）候補が複数ある場合は部屋番号が小さい順を採用
    - 備考に「駐車場(Pxx)→0001」を追記（重複防止）
    """
    to_delete = []

    # 契約者名 → 居室キーの索引（Pで始まらない room のみ）
    by_tenant = {}
    for key, rec in all_recs.items():
        room = (rec.get("room") or "").upper()
        tenant = (rec.get("tenant") or "").strip()
        if not tenant:
            continue
        if not room.startswith("P"):  # 居室のみ
            by_tenant.setdefault(tenant, []).append(key)

    # 居室候補を部屋番号（数字化）で安定ソート
    def room_num_key(k):
        rm = all_recs[k]["room"]
        m = re.sub(r"\D", "", str(rm) or "")
        return int(m) if m else 9999
    for t in by_tenant:
        by_tenant[t].sort(key=room_num_key)

    # Pxx を付替え
    for key, rec in list(all_recs.items()):
        room = (rec.get("room") or "").upper()
        if not room.startswith("P"):
            continue

        tenant = (rec.get("tenant") or "").strip()
        if not tenant:
            logger.info(f"Pxx行 {key} は契約者名が空のため付替え不可（残存）")
            continue

        candidates = by_tenant.get(tenant)
        if not candidates:
            logger.info(f"Pxx行 {key} ({tenant}) は一致する居室が見つからず残存")
            continue

        # 採用先は最有力（最小部屋番号）
        tkey = candidates[0]
        target = all_recs[tkey]
        target_room = target.get("room") or ""

        # 月ごとに駐車料を加算し、備考にメモ
        for mk, mv in (rec.get("monthly") or {}).items():
            dst = target["monthly"].setdefault(
                mk, {"rent":0,"fee":0,"parking":0,"water":0,"reikin":0,"koushin":0,"bikou":""}
            )
            add_p = clean_int(mv.get("parking"))
            if add_p:
                dst["parking"] += add_p
            note = f"駐車場({room})→{target_room}"
            dst["bikou"] = append_note_unique(dst.get("bikou"), note)

        to_delete.append(key)

    for key in to_delete:
        all_recs.pop(key, None)

def most_frequent_amount(values):
    """
    values: 数値リスト
    優先1: 非ゼロの最頻値
    優先2: 0を含めた最頻値
    優先3: 空なら 0
    最頻値が複数ある場合は、その中で最大値を選ぶ
    """
    vals = [clean_int(v) for v in values]

    def pick_mode(nums):
        if not nums:
            return None
        cnt = Counter(nums)
        max_freq = max(cnt.values())
        # 最頻値候補の中で最大値を返す
        candidates = [v for v, f in cnt.items() if f == max_freq]
        return max(candidates)

    # 1. 非ゼロの最頻値
    nonzero = [v for v in vals if v != 0]
    mode = pick_mode(nonzero)
    if mode is not None:
        return mode

    # 2. 全値の最頻値（ゼロ含む）
    mode = pick_mode(vals)
    if mode is not None:
        return mode

    # 3. 空なら 0
    return 0

async def process_files(files):
    tasks = [handle_file(file) for file in files]
    results = await asyncio.gather(*tasks)

    # 1) マージ
    all_recs = {}  # key = (room, tenant)
    for recs in results:
        merge_records(all_recs, recs)

    # 2) Pxx 付替え（契約者名で居室へ）
    fold_parking_Pxx(all_recs)

    # 3) 出力用に並べ替え & 基準額（最頻値）付与
    out = []
    for (room, tenant), rec in all_recs.items():
        def collect(k):
            return [clean_int(v.get(k,0)) for v in rec["monthly"].values()]
        rec["base_rent"]    = most_frequent_amount(collect("rent"))
        rec["base_fee"]     = most_frequent_amount(collect("fee"))
        rec["base_parking"] = most_frequent_amount(collect("parking"))
        rec["base_water"]   = most_frequent_amount(collect("water"))
        out.append(rec)

    # 室番号数値→名前→月最小 でソート
    def room_sort_key(r):
        rm = r["room"]
        if isinstance(rm, str) and rm.upper().startswith("P"):
            num = 9000 + int(re.sub(r"\D","",rm) or 0)  # 駐車は末尾
        else:
            num = int(re.sub(r"\D","",rm) or 0) if re.sub(r"\D","",str(rm)) else 9999
        first_month = sorted(r["monthly"].keys())[0] if r["monthly"] else "9999-99"
        return (num, r["tenant"] or "~", first_month)

    out_sorted = sorted(out, key=room_sort_key)
    months = sorted({m for r in out_sorted for m in r["monthly"].keys()})
    return out_sorted, months

# ========== Excel 生成 ==========
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
    ws.title = xls_clean(property_name) or "入居管理表"

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

    yellow_fill = PatternFill("solid", fgColor="FFF2CC")   # 合計
    pink_fill   = PatternFill("solid", fgColor="F8CBAD")   # 総合計
    thick_side  = Side(style="thick")
    thick_border = Border(left=thick_side, right=thick_side, top=thick_side, bottom=thick_side)
    
    num_months = len(months)
    col_B = 2; col_C = 3; col_D = 4; col_E = 5; col_F = 6; col_G = 7
    col_month_end = 6 + num_months
    col_S = col_month_end + 1
    col_T = col_month_end + 2
    col_U = col_month_end + 3
    col_V = col_month_end + 4
    col_W = col_month_end + 5
    col_X = col_W + 1

    # ---- タイトル & 物件名 ----
    ws.merge_cells(start_row=2, start_column=col_B, end_row=2, end_column=col_W)
    if months:
        start_month = months[0].replace("-", "年") + "月"
        end_month   = months[-1].replace("-", "年") + "月"
        title_val = f"入居管理表 （{start_month}〜{end_month}）"
    else:
        title_val = "入居管理表"
    ws.cell(row=2, column=col_B, value=xls_clean(title_val)).font = Font(size=14, bold=True)
    ws.cell(row=2, column=col_B).alignment = center

    ws.merge_cells(start_row=4, start_column=col_B, end_row=4, end_column=col_C)
    ws.cell(row=4, column=col_B, value="物件名").alignment = center
    ws.merge_cells(start_row=4, start_column=col_D, end_row=4, end_column=col_F)
    ws.cell(row=4, column=col_D, value=xls_clean(property_name or "")).alignment = center
    for c in range(col_B, col_C+1):
        ws.cell(row=4, column=c).border = thick_border
    for c in range(col_D, col_F+1):
        ws.cell(row=4, column=c).border = thick_border

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
        cc.fill = header_fill
        cc.font = bold_font
        cc.alignment = center

    # ---- データ（5行ブロック）----
    row = data_start_row
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
        ws.cell(row=row, column=col_C, value=xls_clean(room)).alignment = center
        ws.cell(row=row, column=col_C).fill = green_fill
        ws.merge_cells(start_row=row+1, start_column=col_C, end_row=row+4, end_column=col_C)
        ws.cell(row=row+1, column=col_C, value=xls_clean(tenant)).alignment = center

        # 科目（D列）と基準額（E列）
        subjects = ["家賃","共益費　","駐車料","水道料","合計"]
        for i, s in enumerate(subjects):
            ws.cell(row=row+i, column=col_D, value=xls_clean(s))
        for i, v in enumerate([base_r, base_f, base_p, base_w]):
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

        # U: 礼金・更新料（5行結合）
        ws.merge_cells(start_row=row, start_column=col_U, end_row=row+4, end_column=col_U)
        cu = ws.cell(row=row, column=col_U, value=reikin_koushin_total); cu.alignment = center_vert; cu.number_format = number_fmt
        # V: 敷金（5行結合）
        ws.merge_cells(start_row=row, start_column=col_V, end_row=row+4, end_column=col_V)
        cv = ws.cell(row=row, column=col_V, value=shikikin); cv.alignment = center_vert; cv.number_format = number_fmt
        # W: 備考（5行結合）
        ws.merge_cells(start_row=row, start_column=col_W, end_row=row+4, end_column=col_W)
        bw = ws.cell(row=row, column=col_W, value=xls_clean(combine_bikou_contract(rec))); bw.alignment = center_vert; bw.font = red_font

        # 罫線・黄色網掛け
        for c in range(col_B, col_W+1):
            for r in range(row, row+5):
                ws.cell(row=r, column=c).border = thin_border
        for c in range(col_B, col_W+1):
            ws.cell(row=row+4, column=c).fill = yellow_fill

        row += 5

    # データ範囲
    first_data_row = data_start_row
    last_data_row  = row - 1

    # ---- 下段「合計」4行 ----
    sum_start = row
    ws.merge_cells(start_row=sum_start, end_row=sum_start+3, start_column=col_B, end_column=col_C)
    ws.cell(row=sum_start, column=col_B, value="合計").alignment = center
    for i, name in enumerate(["家賃","共益費　","駐車料","水道料"]):
        ws.cell(row=sum_start+i, column=col_D, value=xls_clean(name))
    def sumif_range(col_letter):
        return f"{col_letter}${first_data_row}:{col_letter}${last_data_row}"
    for i in range(4):
        r = sum_start + i
        for cidx in range(col_E, col_T+1):
            col_letter = get_column_letter(cidx)
            ws.cell(row=r, column=cidx, value=f"=SUMIF($D${first_data_row}:$D${last_data_row},$D${r},{sumif_range(col_letter)})").number_format = number_fmt
    for cidx in [col_U, col_V]:
        col_letter = get_column_letter(cidx)
        ws.cell(row=sum_start, column=cidx, value=f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})").number_format = number_fmt
        for i in range(1,4):
            ws.cell(row=sum_start+i, column=cidx, value=None)
    for i in range(4):
        ws.cell(row=sum_start+i, column=col_W, value="")

    for c in range(col_B, col_W+1):
        for r in range(sum_start, sum_start+4):
            ws.cell(row=r, column=c).border = thin_border

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
        ws.cell(row=grand_row, column=c).border = thin_border
        ws.cell(row=grand_row, column=c).fill = pink_fill

    # ---- 右外側「確認用」 & 一括チェック式 ----
    ws.cell(row=grand_row-1, column=col_X, value=xls_clean("確認用")).alignment = center
    g_letter = get_column_letter(col_G)
    r_letter = get_column_letter(col_month_end)
    ws.cell(row=grand_row, column=col_X, value=f"=SUM({g_letter}{first_data_row}:{r_letter}{last_data_row})/2").number_format = number_fmt

    # ---- 2行下の「算式確認」行 ----
    check_row = grand_row + 2
    ws.cell(row=check_row, column=col_E, value=xls_clean("算式確認"))
    for cidx in range(col_F, col_T+1):
        col_letter = get_column_letter(cidx)
        ws.cell(row=check_row, column=cidx, value=f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})/2").number_format = number_fmt

    # 備考列の幅（可変）
    ws.column_dimensions[get_column_letter(col_W)].width = max(
        [len(xls_clean(combine_bikou_contract(rec)) or "") for rec in records] + [10]
    ) * 1.6

    # ウィンドウ枠の固定
    try:
        ws.freeze_panes = ws.cell(row=data_start_row, column=last_fixed_col+1)  # "D7"
    except Exception:
        pass

    # 保存
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ========== Streamlit UI ==========
st.set_page_config(page_title="入居管理表アプリ（pdfplumber-only）", layout="wide")
st.title("📊 収支報告書PDFから入居管理表を作成（pdfplumber-only）")

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
        st.info("収支報告書を読み取り中（pdfplumberのみ）...")
        records, months = asyncio.run(process_files(uploaded_files))
        if not records:
            st.error("データが抽出できませんでした。PDFの品質やフォーマットをご確認ください。（pdfplumberのみ運転）")
            st.stop()

        st.info("入居管理表を作成中...")
        excel_data = export_excel(records, months, property_name)
        if months:
            start_month = months[0].replace("-", "年") + "月"
            end_month   = months[-1].replace("-", "年") + "月"
            fn = f"{xls_clean(property_name) or '入居管理表'}（{start_month}〜{end_month}）_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
        else:
            fn = f"{xls_clean(property_name) or '入居管理表'}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"

        st.download_button("入居管理表をダウンロード", data=excel_data,
                           file_name=fn,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success("完了しました。")

