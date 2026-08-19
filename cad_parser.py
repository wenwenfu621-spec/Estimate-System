"""
cad_parser.py - CAD 解析與報表產出模組
Version: v1.8.2_20260819
Description: 支援載入 template.xlsm / template.xlsx 範本檔。
             於 O7 填寫報價日期，並由第 10 列起依序寫入：
             A欄(序項)、B欄(品名/檔名)、P欄(尺寸整合)、Q欄(長)、R欄(寬)、S欄(高)、T欄(單位)。
             完整修復與保護 G/J/M 欄明細公式與第 55 列加總公式 (#REF! 防護)。
"""

import os
import tempfile
from datetime import datetime
from typing import Dict, Any, List
import cadquery as cq
from PIL import Image
import openpyxl


def parse_cad_with_screenshot(file_path: str) -> Dict[str, Any]:
    """
    讀取 CAD 檔案，提取邊界尺寸，並以獨立 try-except 保護截圖生成。
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到檔案：{file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    valid_extensions = ('.step', '.stp', '.igs', '.iges')
    if ext not in valid_extensions:
        raise ValueError(f"不支援的格式 '{ext}'")

    # 1. 核心幾何長寬高解析
    try:
        if ext in ('.step', '.stp'):
            model = cq.importers.importStep(file_path)
        else:
            model = cq.importers.importShape(cq.importers.ImportTypes.IGES, file_path)

        bbox = model.val().BoundingBox()
        x_len = round(bbox.xlen, 2)
        y_len = round(bbox.ylen, 2)
        z_len = round(bbox.zlen, 2)
        dims_sorted: List[float] = sorted([x_len, y_len, z_len], reverse=True)
        
        length_val = dims_sorted[0]
        width_val = dims_sorted[1]
        height_val = dims_sorted[2]
        
        # 尺寸字串格式化為：長*寬*高
        dimensions_str = f"{length_val:.2f}*{width_val:.2f}*{height_val:.2f}"
    except Exception as e:
        return {
            "status": "error",
            "file_name": os.path.basename(file_path),
            "error_message": f"CAD 幾何解析失敗: {str(e)}"
        }

    # 2. 獨立截圖流程 (方案 B 容錯防護)
    img_path = None
    try:
        img_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
        img_path_candidate = img_tmp.name
        img_tmp.close()

        svg_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.svg')
        svg_path = svg_tmp.name
        svg_tmp.close()

        cq.exporters.export(model, svg_path, opt={"projectionDir": (1, 1, 1), "showHidden": False})

        with Image.open(svg_path) as img:
            img_resized = img.resize((400, 300))
            img_resized.save(img_path_candidate, format="PNG")

        img_path = img_path_candidate

        if os.path.exists(svg_path):
            os.remove(svg_path)
    except Exception:
        img_path = None

    return {
        "status": "success",
        "file_name": os.path.basename(file_path),
        "dimensions_str": dimensions_str,
        "length": length_val,
        "width": width_val,
        "height": height_val,
        "unit": "mm",
        "image_path": img_path
    }


def generate_excel_report(parsed_results: List[Dict[str, Any]], output_excel_path: str):
    """
    載入 template.xlsm / template.xlsx 範本檔，寫入解析結果並保護所有計算公式。
    """
    template_candidates = [
        "template.xlsm", "template.xlsx", "template.xls",
        "Template.xlsm", "Template.xlsx", "Template.xls"
    ]
    template_file = None
    for tf in template_candidates:
        if os.path.exists(tf):
            template_file = tf
            break

    is_xlsm = template_file and template_file.lower().endswith('.xlsm')

    if template_file:
        wb = openpyxl.load_workbook(template_file, data_only=False, keep_vba=is_xlsm)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "報價單"

    # 1. 填入當天報價日期於 O7 欄位
    today_str = datetime.now().strftime("%Y/%m/%d")
    ws['O7'] = today_str

    # 2. 從第 10 列起逐筆填入資料與拆解尺寸
    start_row = 10
    for idx, item in enumerate(parsed_results):
        row_num = start_row + idx
        
        # A 欄: 序項 (1, 2, 3...)
        ws.cell(row=row_num, column=1, value=idx + 1)
        
        # B 欄: 品名 (上傳的 CAD 檔名)
        ws.cell(row=row_num, column=2, value=item.get("file_name", ""))
        
        # P 欄: 尺寸整合字串 (長*寬*高)
        ws.cell(row=row_num, column=16, value=item.get("dimensions_str", ""))
        
        # Q 欄: 長 (Length)
        if "length" in item:
            ws.cell(row=row_num, column=17, value=item["length"])
            
        # R 欄: 寬 (Width)
        if "width" in item:
            ws.cell(row=row_num, column=18, value=item["width"])
            
        # S 欄: 高 (Height)
        if "height" in item:
            ws.cell(row=row_num, column=19, value=item["height"])
            
        # T 欄: 單位 (mm)
        ws.cell(row=row_num, column=20, value=item.get("unit", "mm"))

        # 明細列自動金額計算公式防護 (數量 * 單價)
        ws.cell(row=row_num, column=7, value=f"=E{row_num}*F{row_num}")  # G欄: 原型金額
        ws.cell(row=row_num, column=10, value=f"=H{row_num}*I{row_num}") # J欄: 矽膠模具金額
        ws.cell(row=row_num, column=13, value=f"=K{row_num}*L{row_num}") # M欄: 注型金額

    # 3. 第 55 列加總公式防護 (徹底修復與防止 #REF! 錯誤)
    ws['G55'] = "=SUM(G10:G54)"
    ws['J55'] = "=SUM(J10:J54)"
    ws['M55'] = "=SUM(M10:M54)"
    ws['O55'] = "=G55+J55+M55"

    wb.save(output_excel_path)
