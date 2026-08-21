"""
pdf_2d_parser.py - 輕量化 PDF 2D 尺寸標註與上下文解析模組
Version: v2.8.5_20260821
Description: 使用 pypdf 提取文字與座標，透過正規表達式捕捉尺寸與公差 (如 69.7±0.3)，
             自動過濾 Notes、Adhesive、厚度等干擾項，提供候選尺寸並支援 Fallback。
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
    """解析 PDF 取得最大外觀尺寸 (Length x Width)"""
    if not os.path.exists(file_path):
        return {"status": "error", "error_message": "找不到 PDF 檔案"}
    
    if not HAS_PYPDF:
        return {"status": "error", "error_message": "系統缺少 pypdf 套件，無法解析 PDF"}

    text_elements = []
    try:
        reader = PdfReader(file_path)
        for page_idx, page in enumerate(reader.pages):
            def visitor_body(text, cm, tm, font_dict, font_size):
                if text and text.strip():
                    # tm[4], tm[5] 為 x, y 座標
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
        return {"status": "error", "error_message": "PDF 內無可讀取文字（可能是掃描圖片型 PDF）", "status_code": "needs_manual_input"}

    # 排除關鍵字（Notes / 膠水 / 厚度 / R角 / 直徑等）
    exclude_keywords = [
        "THICKNESS", "ADHESIVE", "R", "RADIUS", "Ø", "DIA", 
        "HOLE", "PITCH", "TYP", "REF", "SCALE", "DATE", "REV", 
        "NOTE", "MATERIAL", "HARDNESS", "MM", "INCH"
    ]

    candidates = []

    # 匹配模式：支援 69.7, 69.7±0.3, 69.7 ± 0.3, 69.7+0.3/-0.2 等
    dim_pattern = re.compile(r'^([0-9]+(?:\.[0-9]+)?)\s*(?:[±±+-]\s*[0-9]+(?:\.[0-9]+)?(?:/[0-9]+(?:\.[0-9]+)?)?)?$')

    for el in text_elements:
        txt = el["text"].upper()
        # 檢查是否含有排除關鍵字
        if any(kw in txt for kw in exclude_keywords):
            continue

        match = dim_pattern.match(txt)
        if match:
            try:
                val = float(match.group(1))
                # 合理的 Mylar 零件外觀尺寸範圍 (例如 1 mm 到 2000 mm)
                if 1.0 <= val <= 3000.0:
                    candidates.append({
                        "nominal": val,
                        "raw_text": el["text"],
                        "x": el["x"],
                        "y": el["y"],
                        "page": el["page"]
                    })
            except Exception:
                pass

    if len(candidates) < 2:
        return {
            "status": "error", 
            "error_message": "無法從 PDF 中可靠辨識出足夠的外觀尺寸標註", 
            "status_code": "needs_manual_input"
        }

    # 簡單候選排序：依數值大小分組或取前兩個合理的整體尺寸候選（排除過小的數值如 2.8、0.15 等）
    valid_vals = [c["nominal"] for c in candidates if c["nominal"] > 5.0]
    
    # 若有效數值大於等於 2 個，取最大的兩個作為 Length 與 Width
    unique_sorted_vals = sorted(list(set(valid_vals)), reverse=True)

    if len(unique_sorted_vals) >= 2:
        l = unique_sorted_vals[0]
        w = unique_sorted_vals[1]
        length = max(l, w)
        width = min(l, w)
        return {
            "status": "success",
            "file_type": "PDF",
            "length": length,
            "width": width,
            "dimensions_str": f"{length}*{width}",
            "dimension_source": "PDF Dimension"
        }

    return {
        "status": "error",
        "error_message": "PDF 尺寸候選不足或無法判定整體外框",
        "status_code": "needs_manual_input"
    }
