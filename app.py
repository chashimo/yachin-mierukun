# 非同期処理対応の完全修正済みスクリプト（AsyncOpenAI 使用）

import streamlit as st
import os
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

# ======= ログ設定 =======
if not logging.getLogger().hasHandlers():
    log_filename = f"app_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format='[%(asctime)s] %(levelname)s: %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
logger = logging.getLogger(__name__)

# ========== OpenAI 非同期クライアント設定 ==========
client = AsyncOpenAI(api_key=st.secrets["OPENAI_API_KEY"])

async def call_openai_vision_async(base64_images, text_context, default_month_id):
    image_parts = [{"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}} for b64 in base64_images]

    messages = [
        {"role": "system", "content":
            "あなたは不動産管理アシスタントです。収支報告書のPDF画像とテキストから、"
            "各部屋の入居情報を以下のJSON構造に変換してください。"
        },
        {"role": "user", "content": [
            *image_parts,
            {"type": "text", "text":
                f"以下はこのPDFからテキスト抽出した内容です（文字認識補正の参考にしてください）:\n{text_context}\n\n"
                f"この収支報告書には複数の月（たとえば {default_month_id} やその前後の月）が記載されている可能性があります。\n"
                f"表内に出現するすべての「年／月」に対応する情報を抽出してください。\n\n"
                "出力形式は次のようにしてください：\n"
                "{\n"
                "  \"0101\": {\n"
                "    \"name\": \"入居者名\",\n"
                "    \"base_rent\": 家賃の基準額,\n"
                "    \"base_kyoueihi\": 共益費の基準額,\n"
                "    \"monthly\": {\n"
                "        \"2024-12\": {\"rent\": 実家賃, \"kyoueihi\": 実共益費, \"bikou\": \"備考\"},\n"
                "        \"2025-01\": {\"rent\": 実家賃, \"kyoueihi\": 実共益費, \"bikou\": \"備考\"},\n"
                "        ...\n"
                "    },\n"
                "    \"reikin\": 礼金,\n"
                "    \"shikikin\": 敷金,\n"
                "    \"koushinryo\": 更新料\n"
                "  },\n"
                "  ...\n"
                "}\n\n"
                "空室の部屋は出力しないでください。\n"
                "出力は ```json や ``` で囲まず、プレーンな JSON オブジェクトのみを返してください。\n"
                "すべての数値（家賃、共益費、礼金など）はカンマなしの整数値で出力してください（例：61000）。"
            }
        ]}
    ]

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        temperature=0.0,
        max_tokens=4096
    )
    return response.choices[0].message.content

# ========== ユーティリティ ==========
def convert_pdf_to_images(pdf_bytes, dpi=200):
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in pdf:
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    return images

def convert_image_to_base64(image):
    buf = io.BytesIO()
    image.save(buf, format="JPEG")
    return base64.b64encode(buf.getvalue()).decode("utf-8")

def extract_text_with_pdfplumber(pdf_bytes):
    texts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or "")
    return "\n".join(texts)

def extract_month_from_filename(filename: str) -> str:
    match = re.search(r"(\d{4})年(\d{1,2})月", filename)
    if match:
        return f"{match.group(1)}-{match.group(2).zfill(2)}"
    return "unknown"

# ========== 各ファイル処理 ==========
async def handle_file(file, max_attempts=3):
    file_name = file.name
    logger.info(f"開始: {file_name}")
    default_month_id = extract_month_from_filename(file_name)
    file_bytes = file.read()
    images = convert_pdf_to_images(file_bytes)
    base64_images = [convert_image_to_base64(img) for img in images]
    text_context = extract_text_with_pdfplumber(file_bytes)

    for attempt in range(1, max_attempts + 1):
        try:
            json_str = await call_openai_vision_async(base64_images, text_context, default_month_id)
            json_str_clean = json_str.strip().removeprefix("```json").removeprefix("```").removesuffix("```")
            if not json_str_clean.strip().startswith("{"):
                raise ValueError("OpenAIの出力がJSON形式ではありません")
            partial = json.loads(json_str_clean)
            logger.info(f"{file_name}: JSON解析成功")
            return partial
        except Exception as e:
            logger.warning(f"{file_name}: JSON解析失敗（試行{attempt}回目）: {e}")
            if attempt == max_attempts:
                return file_name, None

# ========== 全ファイル並列処理 ==========
async def process_files(files):
    tasks = [handle_file(file) for file in files]
    results = await asyncio.gather(*tasks)
    all_data = {}
    for result in results:
        if isinstance(result, tuple):  # (filename, None)
            st.warning(f"{result[0]} の出力がJSONとして解釈できませんでした。")
            continue
        for room_id, info in result.items():
            if room_id not in all_data:
                all_data[room_id] = info
            else:
                for key in ["name", "reikin", "shikikin", "koushinryo"]:
                    if info.get(key):
                        all_data[room_id][key] = info[key]
                all_data[room_id]["monthly"].update(info.get("monthly", {}))
    return all_data

# ========== Excel生成 ==========
def combine_bikou(info):
    bikou_set = set()
    for month_data in info.get("monthly", {}).values():
        b = month_data.get("bikou")
        if b and isinstance(b, str) and b.strip():
            bikou_set.add(b.strip())
    return ", ".join(sorted(bikou_set))

def export_excel(all_data, property_name):
    wb = Workbook()
    ws = wb.active
    ws.title = property_name

    months = sorted(set(m for info in all_data.values() for m in info.get("monthly", {})))
    num_months = len(months)

    header_fill = PatternFill("solid", fgColor="BDD7EE")
    green_fill = PatternFill("solid", fgColor="CCFFCC")
    gray_fill = PatternFill("solid", fgColor="DDDDDD")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    center_vert = Alignment(vertical="center", wrap_text=True)
    bold_font = Font(bold=True)
    red_font = Font(color="9C0000")
    number_format = "#,##0"
    thin_border = Border(*[Side(style='thin')] * 4)

    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=5 + num_months + 4)
    start_month = months[0].replace("-", "年") + "月"
    end_month = months[-1].replace("-", "年") + "月"
    ws["B2"] = f"{property_name} 入居管理表 （{start_month}〜{end_month}）"
    ws["B2"].font = Font(size=14, bold=True)
    ws["B2"].alignment = center

    ws.merge_cells("B5:C5")
    ws["B5"] = "賃借人"
    ws.merge_cells("D5:E5")
    ws["D5"] = "基準額"
    ws["F5"] = "期首 未収/前受"

    for i, m in enumerate(months):
        ws.cell(row=5, column=7 + i, value=f"{int(m[5:])}月")

    labels = ["合計", "期末 未収/前受", "礼金・更新料", "敷金", "備考"]
    for i, label in enumerate(labels):
        ws.cell(row=5, column=7 + num_months + i, value=label)

    col_bikou = 7 + num_months + 4
    for col in range(2, col_bikou + 1):
        cell = ws.cell(row=5, column=col)
        cell.fill = header_fill
        cell.font = bold_font
        cell.alignment = center

    row_idx = 6
    for room_id in sorted(all_data):
        info = all_data[room_id]
        name = info.get("name", "")
        rent_base = max([m.get("rent", 0) for m in info.get("monthly", {}).values()])
        fee_base = max([m.get("kyoueihi", 0) for m in info.get("monthly", {}).values()])
        total_base = rent_base + fee_base

        ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx + 2, end_column=2)
        cell = ws.cell(row=row_idx, column=2, value="室番号")
        cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

        ws[f"C{row_idx}"] = room_id
        ws[f"C{row_idx}"].alignment = center
        ws[f"C{row_idx}"].fill = green_fill

        ws.merge_cells(start_row=row_idx + 1, start_column=3, end_row=row_idx + 2, end_column=3)
        ws[f"C{row_idx + 1}"] = name
        ws[f"C{row_idx + 1}"].alignment = center

        ws[f"D{row_idx}"] = "家賃"
        ws[f"D{row_idx + 1}"] = "共益費"
        ws[f"D{row_idx + 2}"] = "合計"

        for r, val in enumerate([rent_base, fee_base, total_base]):
            cell = ws.cell(row=row_idx + r, column=5, value=val)
            cell.number_format = number_format

        rent_sum = fee_sum = 0
        for i, month in enumerate(months):
            rent = info.get("monthly", {}).get(month, {}).get("rent", 0)
            fee = info.get("monthly", {}).get(month, {}).get("kyoueihi", 0)
            total = rent + fee
            rent_sum += rent
            fee_sum += fee
            for r, val in enumerate([rent, fee, total]):
                cell = ws.cell(row=row_idx + r, column=7 + i, value=val)
                cell.number_format = number_format

        for r, val in enumerate([rent_sum, fee_sum, rent_sum + fee_sum]):
            cell = ws.cell(row=row_idx + r, column=7 + num_months)
            cell.number_format = number_format

        ws.merge_cells(start_row=row_idx, start_column=col_bikou - 2, end_row=row_idx + 2, end_column=col_bikou - 2)
        ws.merge_cells(start_row=row_idx, start_column=col_bikou - 1, end_row=row_idx + 2, end_column=col_bikou - 1)
        ws.merge_cells(start_row=row_idx, start_column=col_bikou, end_row=row_idx + 2, end_column=col_bikou)

        ws.cell(row=row_idx, column=col_bikou - 2, value=info.get("reikin", "")).alignment = center_vert
        ws.cell(row=row_idx, column=col_bikou - 1, value=info.get("shikikin", "")).alignment = center_vert
        bikou_cell = ws.cell(row=row_idx, column=col_bikou, value=combine_bikou(info))
        bikou_cell.alignment = center_vert
        bikou_cell.font = red_font

        for c in range(2, col_bikou + 1):
            for r in range(row_idx, row_idx + 3):
                ws.cell(row=r, column=c).border = thin_border
        for c in range(2, col_bikou + 1):
            ws.cell(row=row_idx + 2, column=c).fill = gray_fill

        row_idx += 3

    max_len = max(len(combine_bikou(info)) for info in all_data.values())
    ws.column_dimensions[get_column_letter(col_bikou)].width = max_len * 1.5

    out_file = io.BytesIO()
    wb.save(out_file)
    return out_file.getvalue(), start_month, end_month

# ========== Streamlit UI ==========
st.set_page_config(page_title="入居管理表アプリ", layout="wide")
st.title("📊 収支報告書PDFから入居管理表を作成")

PASSWORD = st.secrets["APP_PASSWORD"]
pw = st.text_input("パスワードを入力してください", type="password")
if pw != PASSWORD:
    st.warning("正しいパスワードを入力してください。")
    st.stop()

property_name = st.text_input("物件名（例：ジーメゾン入間東藤沢）", value="")
uploaded_files = st.file_uploader("収支報告書PDFを最大12ファイルまでアップロードしてください", type="pdf", accept_multiple_files=True)

if uploaded_files and st.button("入居管理表を作成"):
    if len(uploaded_files) > 12:
        st.warning("アップロードできるのは最大12ファイルまでです。")
    else:
        st.info("収支報告書を読み取り中...")
        all_data = asyncio.run(process_files(uploaded_files))
        st.info("入居管理表を作成中...")
        excel_data, start_month, end_month = export_excel(all_data, property_name)
        filename = f"{property_name}_入居管理表（{start_month}〜{end_month}）_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
        st.download_button("入居管理表をダウンロード", data=excel_data, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.info("新規に入居管理表を作成する場合は、ブラウザのリロードボタンを押してください。")
