"""
pdf_2d_parser.py - 輕量化 PDF 2D 尺寸標註提取模組
Version: v2.8.4_20260821
Description: 使用 pypdf 提取純文字，利用正規表達式搜尋長寬標註，
             避免直接取最大數字陷阱，失敗時回傳錯誤供手動 Fallback。
"""

import os
import re
from typing import Dict, Any

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    try:
        from PyPDF2 import PdfReader
        HAS_PYPDF = True
    except ImportError:
        HAS_PYPDF = False


def parse_pdf_2d(file_path: str) -> Dict[str, Any]:
    """解析 PDF 取得最大外觀尺寸 (Length x Width)"""
    if not os.path.exists(file_path):
        return {"status": "error", "error_message": "找不到 PDF 檔案"}
    
    if not HAS_PYPDF:
        return {"status": "error", "error_message": "系統缺少 pypdf 套件，無法解析 PDF"}

    extracted_text = ""
    try:
        reader = PdfReader(file_path)
        for page in reader.pages:
            txt = page.extract_text()
            if txt:
                extracted_text += txt + "\n"
    except Exception as e:
        return {"status": "error", "error_message": f"PDF 讀取例外: {str(e)}"}

    if not extracted_text.strip():
        return {"status": "error", "error_message": "PDF 內無可讀取文字（可能是掃描圖片型 PDF）"}

    found_pairs = []
    # 支援格式如 325x186, 325*186, 325 X 186 等
    patterns = [
        r'(?:overall|dimension|length|width|長|寬|外形|尺寸)?\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)\s*[xX*×]\s*([0-9]+(?:\.[0-9]+)?)',
        r'([0-9]+(?:\.[0-9]+)?)\s*[xX*×]\s*([0-9]+(?:\.[0-9]+)?)'
    ]

    for pat in patterns:
        matches = re.findall(pat, extracted_text, re.IGNORECASE)
        for m in matches:
            try:
                v1, v2 = float(m[0]), float(m[1])
                if 1.0 <= v1 <= 5000 and 1.0 <= v2 <= 5000:
                    found_pairs.append((v1, v2))
            except Exception:
                pass

    if found_pairs:
        l, w = found_pairs[0]
        length = max(l, w)
        width = min(l, w)
        return {
            "status": "success",
            "file_type": "PDF",
            "length": length,
            "width": width,
            "dimensions_str": f"{length}*{width}",
            "source": "PDF Dimension"
        }

    return {
        "status": "error",
        "error_message": "無法從 PDF 文字中可靠辨識出外觀尺寸標註"
    }
