"""
cad_parser.py - CAD 解析與報表產出模組
Version: v1.5.1_20260818
Description: 支援載入 template.xlsm / template.xlsx 範本檔。
             於 O7 填寫報價日期，並由第 10 列起依序寫入：
             A欄(序項)、B欄(品名/檔名)、P欄(尺寸 長*寬*高)、T欄(單位)，
             輸出為標準 .xlsx 格式以避免 Excel 跳出檔案損毀警告。
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
        
        # 尺寸字串格式化為：長*寬*高
        dimensions_str = f"{dims_sorted[0]:.2f}*{dims_sorted[1]:.2f}*{dims_sorted[2]:.2f}"
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
        "unit": "mm",
        "image_path": img_path
    }


def generate_excel_report(parsed_results: List[Dict[str, Any]], output_excel_path: str):
    """
    載入 template.xlsm / template.xlsx 範本檔，寫入解析結果並保留原公式。
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

    if template_file:
        # keep_vba=True 可完整保留巨集與公式結構
        wb = openpyxl.load_workbook(template_file, keep_vba=template_file.endswith('.xlsm'))
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "報價單"

    # 1. 填入當天報價日期於 O7 欄位
    today_str = datetime.now().strftime("%Y/%m/%d")
    ws['O7'] = today_str

    # 2. 從第 10 列起逐筆填入資料
    start_row = 10
    for idx, item in enumerate(parsed_results):
        row_num = start_row + idx
        
        # A 欄: 序項 (1, 2, 3...)
        ws.cell(row=row_num, column=1, value=idx + 1)
        
        # B 欄: 品名 (上傳的 CAD 檔名)
        ws.cell(row=row_num, column=2, value=item.get("file_name", ""))
        
        # P 欄: 尺寸 (長*寬*高)
        ws.cell(row=row_num, column=16, value=item.get("dimensions_str", ""))
        
        # T 欄: 單位 (mm)
        ws.cell(row=row_num, column=20, value=item.get("unit", "mm"))

    wb.save(output_excel_path)
