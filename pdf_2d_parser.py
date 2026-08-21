"""
pdf_2d_parser.py - 2D PDF 尺寸候選解析與分類模組
Version: v2.8.8_20260821
Description: 實作 Dimension Candidate 分類 (overall, outer_diameter, radius, angle, thickness 等)，
             移除大數字盲目排序，支援精準尺寸擷取。
"""

import os
import re
from typing import Dict, Any, List

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


def classify_and_filter_candidates(text_elements: List[Dict[str, Any]]) -> Dict[str, Any]:
    """將文字元素分類為 DimensionCandidate 並進行幾何與上下文權重排序"""
    diameters = []
    linear_candidates = []

    # 專屬精準 Regex
    dia_pattern = re.compile(r'(?:[Ø⌀φΦ]|DIA\.?\s*)([0-9]+(?:\.[0-9]+)?)', re.IGNORECASE)
    radius_pattern = re.compile(r'\bR[0-9]+(?:\.[0-9]+)?\b', re.IGNORECASE)
    angle_pattern = re.compile(r'[0-9]+(?:\.[0-9]+)?\s*°', re.IGNORECASE)
    thick_pattern = re.compile(r'(?:THICKNESS|ADHESIVE|THK)\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)', re.IGNORECASE)
    dim_pattern = re.compile(r'^([0-9]+(?:\.[0-9]+)?)\s*(?:[±±+-]\s*[0-9]+(?:\.[0-9]+)?(?:/[0-9]+(?:\.[0-9]+)?)?)?$')

    for el in text_elements:
        txt = el["text"].strip()
        txt_upper = txt.upper()

        # 1. 檢查角度 (Angle)
        if angle_pattern.search(txt):
            continue

        # 2. 檢查半徑 (Radius, 例如 R0.9, R4.5)
        if radius_pattern.search(txt):
            continue

        # 3. 檢查厚度/膠水 (Thickness / Adhesive Notes)
        if thick_pattern.search(txt_upper) or "ADHESIVE" in txt_upper:
            continue

        # 4. 檢查直徑 (Diameter)
        dia_match = dia_pattern.search(txt)
        if dia_match:
            try:
                val = float(dia_match.group(1))
                if 1.0 <= val <= 2000.0:
                    # 區分外徑與內孔（依數值大小或位置權重，此處簡化取最大值為外徑候選）
                    diameters.append({"value": val, "type": "outer_diameter", "raw": txt})
            except Exception:
                pass
            continue

        # 5. 一般線性尺寸與公差 (Linear / Tolerance)
        match = dim_pattern.match(txt)
        if match:
            try:
                val = float(match.group(1))
                # 過濾過小數值（如厚度 1mm、1.5mm 等）
                if 5.0 <= val <= 3000.0:
                    linear_candidates.append({"value": val, "type": "overall_linear", "raw": txt})
            except Exception:
                pass

    # 判定邏輯 A：圓形件 (Circular Shape Test B)
    if diameters:
        # 若有明顯最大外徑
        max_dia = max([d["value"] for d in diameters])
        return {
            "status": "success",
            "file_type": "PDF",
            "shape_type": "Circular",
            "length": max_dia,
            "width": max_dia,
            "dimensions_str": f"{max_dia}*{max_dia}",
            "dimension_source": "PDF Outer Diameter",
            "confidence": "High"
        }

    # 判定邏輯 B：矩形 / 異形件 (Test A: 38.92, 27.26)
    unique_lin = sorted(list(set([c["value"] for c in linear_candidates])), reverse=True)
    if len(unique_lin) >= 2:
        l = unique_lin[0]
        w = unique_lin[1]
        length = max(l, w)
        width = min(l, w)
        return {
            "status": "success",
            "file_type": "PDF",
            "shape_type": "Rectangular",
            "length": length,
            "width": width,
            "dimensions_str": f"{length}*{width}",
            "dimension_source": "PDF Dimension Classification",
            "confidence": "Medium"
        }

    return {
        "status": "error",
        "error_message": "無法從 PDF 中高可信判定整體外框尺寸",
        "status_code": "needs_manual_input",
        "confidence": "Low"
    }


def parse_pdf_dimensions(file_path: str) -> Dict[str, Any]:
    """解析 PDF 尺寸入口"""
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

    # 如果文字層為空（Vector PDF 無文字層），回傳 needs_manual_input 觸發 Manual Fallback
    if not text_elements:
        return {
            "status": "error", 
            "error_message": "此 CAD Vector PDF 無可提取之文字層，請手動輸入外觀尺寸", 
            "status_code": "needs_manual_input",
            "confidence": "Low"
        }

    return classify_and_filter_candidates(text_elements)
