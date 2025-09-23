#!/usr/bin/env python3
# rasterize_pdf_to_imagepdf.py
import io, sys, fitz  # PyMuPDF
from PIL import Image

def rasterize(in_path: str, out_path: str, dpi: int = 200):
    src = fitz.open(in_path)
    dst = fitz.open()
    scale = dpi / 72.0
    mat = fitz.Matrix(scale, scale)

    for p in src:
        pix = p.get_pixmap(matrix=mat, alpha=False)  # ラスタライズ
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        # 新しいPDFページに画像を貼る
        page = dst.new_page(width=pix.width, height=pix.height)
        img_buf = io.BytesIO()
        img.save(img_buf, format="JPEG", quality=90, optimize=True)
        rect = fitz.Rect(0, 0, pix.width, pix.height)
        page.insert_image(rect, stream=img_buf.getvalue())

    dst.save(out_path)
    dst.close(); src.close()

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python rasterize_pdf_to_imagepdf.py input.pdf output.pdf [dpi]")
        sys.exit(2)
    dpi = int(sys.argv[3]) if len(sys.argv) >= 4 else 200
    rasterize(sys.argv[1], sys.argv[2], dpi)

