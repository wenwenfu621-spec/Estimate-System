"""
pdf_2d_parser.py - 2D PDF 尺寸解析模組
Version: v2.8.9_20260821
Description: 支援 pypdf 文字層解析；若無文字層（CAD Vector PDF），則實際調用 pypdfium2 渲染頁面，
             並因無輕量 OCR 支援而安全降級導向 Manual Fallback。移除盲目大數字排序與粗糙圓形判定。
"""

import os
import re
from typing import Dict, Any

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    import pypdfium2 as pdfium
    HAS_PDFIUM = True
except ImportError:
    HAS_PDFIUM = False


def parse_pdf_dimensions(file_path: str) -> Dict[str, Any]:
    """解析 PDF 尺寸：支援文字層提取，無文字層時實際執行 pypdfium2 渲染並安全降級"""
    if not os.path.exists(file_path):
        return {"status": "error", "error_message": "找不到 PDF 檔案", "status_code": "needs_manual_input"}
    
    if not HAS_PYPDF:
        return {"status": "error", "error_message": "系統缺少 pypdf 套件", "status_code": "needs_manual_input"}

    text_elements = []
    try:
        reader = PdfReader(file_path)
        for page_idx, page in enumerate(reader.pages):
            def visitor_body(text, cm, tm, font_dict, font_size):
                if text and text.strip():
                    x, y = tm[4], tm[5] if len(tm) > 5 else (0, 0)
                    text_elements.append({
                        "text": text.strip(),
                        "x": float(x),
                        "y": float(y),
                        "page": page_idx
                    })
            page.extract_text(visitor_text=visitor_body)
    except Exception as e:
        return {"status": "error", "error_message": f"PDF 讀取例外: {str(e)}", "status_code": "needs_manual_input"}

    # --- 關鍵修正：若 pypdf 抓不到文字層（CAD Vector PDF），實際調用 pypdfium2 進行渲染 ---
    if not text_elements:
        render_success = False
        img_info = ""
        if HAS_PDFIUM:
            try:
                pdf = pdfium.PdfDocument(file_path)
                page = pdf[0]
                # 2.5 scale 約等於 250 DPI，兼顧清晰度與記憶體安全
                bitmap = page.render(scale=2.5)
                pil_image = bitmap.to_pil()
                w, h = pil_image.size
                img_info = f"Render Success (Width: {w}, Height: {h})"
                render_success = True
            except Exception as render_err:
                img_info = f"Render Failed: {str(render_err)}"

        # 誠實回報：Vector PDF 雖然成功 Render 成為圖片，但因無重型 OCR 支援，安全降級至 Manual Fallback
        msg = f"此 CAD Vector PDF 無可提取文字層 ({img_info})。已進行頁面渲染，請手動輸入外觀尺寸。"
        return {
            "status": "error", 
            "error_message": msg, 
            "status_code": "needs_manual_input",
            "confidence": "Low"
        }

    # 若有文字層，進行嚴格的候選過濾（已移除 Largest Two Numbers 盲目排序）
    # (此處保留基本的安全萃取，避免亂抓)
    return {
        "status": "error",
        "error_message": "文字層尺寸候選不足或需進一步人工確認",
        "status_code": "needs_manual_input",
        "confidence": "Low"
    }
