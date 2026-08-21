"""
pdf_2d_parser.py - 輕量化 PDF 2D 尺寸標註解析模組
Version: v2.8.7_20260821
Description: 支援 pypdf 文字層萃取；若遇無文字層之 CAD 向量 PDF，則安全降級回傳 needs_manual_input。
"""

import os
import re
from typing import Dict, Any

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False


def parse_pdf_dimensions(file_path: str) -> Dict[str, Any]:
    """解析 PDF 取得尺寸，無文字層時安全降級要求手動輸入"""
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

    # 如果沒有文字層（例如全為 Vector Path 且 pypdf 抓不到字串），安全降級交由人工輸入
    if not text_elements:
        return {
            "status": "error", 
            "error_message": "此 PDF 無可提取之文字層（CAD 向量圖面），請手動輸入外觀尺寸", 
            "status_code": "needs_manual_input"
        }

    exclude_keywords = [
        "THICKNESS", "ADHESIVE", "R", "RADIUS", "HOLE", "PITCH", 
        "TYP", "REF", "SCALE", "DATE", "REV", "NOTE", "MATERIAL", "HARDNESS", "MM", "INCH"
    ]

    diameters = []
    standard_dims = []

    dia_pattern = re.compile(r'(?:[Ø⌀φΦ]|DIA\.?\s*)([0-9]+(?:\.[0-9]+)?)', re.IGNORECASE)
    dim_pattern = re.compile(r'^([0-9]+(?:\.[0-9]+)?)\s*(?:[±±+-]\s*[0-9]+(?:\.[0-9]+)?(?:/[0-9]+(?:\.[0-9]+)?)?)?$')

    for el in text_elements:
        txt = el["text"].upper()
        
        dia_match = dia_pattern.search(txt)
        if dia_match:
            try:
                val = float(dia_match.group(1))
                if 1.0 <= val <= 2000.0:
                    diameters.append(val)
            except Exception:
                pass
            continue

        if any(kw in txt for kw in exclude_keywords):
            continue

        match = dim_pattern.match(txt)
        if match:
            try:
                val = float(match.group(1))
                if 5.0 <= val <= 3000.0:
                    standard_dims.append(val)
            except Exception:
                pass

    if diameters:
        max_dia = max(diameters)
        return {
            "status": "success",
            "file_type": "PDF",
            "shape_type": "Circular",
            "length": max_dia,
            "width": max_dia,
            "dimensions_str": f"{max_dia}*{max_dia}",
            "dimension_source": "PDF Outer Diameter"
        }

    unique_dims = sorted(list(set(standard_dims)), reverse=True)
    if len(unique_dims) >= 2:
        l = unique_dims[0]
        w = unique_dims[1]
        length = max(l, w)
        width = min(l, w)
        return {
            "status": "success",
            "file_type": "PDF",
            "shape_type": "Rectangular",
            "length": length,
            "width": width,
            "dimensions_str": f"{length}*{width}",
            "dimension_source": "PDF Dimension"
        }

    return {
        "status": "error",
        "error_message": "無法從 PDF 文字中可靠辨識出整體外框尺寸",
        "status_code": "needs_manual_input"
    }
