"""
pdf_2d_parser.py - 輕量化 PDF 2D 尺寸標註與上下文解析模組
Version: v2.8.6_20260821
Description: 使用 pypdf 提取文字與座標，支援公差綁定、直徑/圓形件識別與 Notes 降權。
"""

import os
import re
from typing import Dict, Any, List

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False


def parse_pdf_dimensions(file_path: str) -> Dict[str, Any]:
    """解析 PDF 取得最大外觀尺寸或外徑 (Length x Width / Circular)"""
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

    if not text_elements:
        return {"status": "error", "error_message": "PDF 無文字層（掃描型 PDF）", "status_code": "needs_manual_input"}

    # 排除純 Notes / 材質 / 單位等關鍵字
    exclude_keywords = [
        "THICKNESS", "ADHESIVE", "R", "RADIUS", "HOLE", "PITCH", 
        "TYP", "REF", "SCALE", "DATE", "REV", "NOTE", "MATERIAL", "HARDNESS", "MM", "INCH"
    ]

    diameters = []
    standard_dims = []

    # 匹配直徑 (如 Ø17.1, DIA 17.1)
    dia_pattern = re.compile(r'(?:[Ø⌀φΦ]|DIA\.?\s*)([0-9]+(?:\.[0-9]+)?)', re.IGNORECASE)
    # 匹配一般尺寸與公差 (如 69.7, 69.7±0.3, 38.92)
    dim_pattern = re.compile(r'^([0-9]+(?:\.[0-9]+)?)\s*(?:[±±+-]\s*[0-9]+(?:\.[0-9]+)?(?:/[0-9]+(?:\.[0-9]+)?)?)?$')

    for el in text_elements:
        txt = el["text"].upper()
        
        # 檢查是否為直徑
        dia_match = dia_pattern.search(txt)
        if dia_match:
            try:
                val = float(dia_match.group(1))
                if 1.0 <= val <= 2000.0:
                    diameters.append(val)
            except Exception:
                pass
            continue

        # 檢查是否含有排除關鍵字
        if any(kw in txt for kw in exclude_keywords):
            continue

        match = dim_pattern.match(txt)
        if match:
            try:
                val = float(match.group(1))
                # 排除過小數值（如厚度 1mm 或 2.8mm 等局部特徵）
                if 5.0 <= val <= 3000.0:
                    standard_dims.append(val)
            except Exception:
                pass

    # 判定邏輯 1：若存在明顯的最大外徑（針對圓形件如 Ø17.1 與內孔 Ø6.x）
    if diameters:
        # 取最大的直徑作為外徑 (Outer Diameter)
        max_dia = max(diameters)
        # 若有其他較小的直徑（如孔徑），確保 max_dia 是顯著的外or 主尺寸
        return {
            "status": "success",
            "file_type": "PDF",
            "shape_type": "Circular",
            "length": max_dia,
            "width": max_dia,
            "dimensions_str": f"{max_dia}*{max_dia}",
            "dimension_source": "PDF Outer Diameter"
        }

    # 判定邏輯 2：矩形或一般異形件，取前兩個最大的獨立尺寸作為 Length 與 Width (如 Test A: 38.92, 27.26; Test C: 69.7, 12.8)
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
        "error_message": "無法從 PDF 中可靠辨識出整體外框尺寸",
        "status_code": "needs_manual_input"
    }
