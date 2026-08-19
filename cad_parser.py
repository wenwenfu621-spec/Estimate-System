"""
cad_parser.py - CAD 解析與報表產出模組
Version: v1.9.0_20260819
Description: 支援載入 template.xlsm / template.xlsx 範本檔。
             採用方案 A：底層雙重相容引擎解析 .iges / .igs / .step / .stp 檔案。
             精準寫入 A, B, P, Q, R, S, T 欄數據，使用 set_cell_value_safe 保護合計列公式。
"""

import os
import tempfile
from datetime import datetime
from typing import Dict, Any, List
import cadquery as cq
from PIL import Image
import openpyxl


def set_cell_value_safe(ws, row: int, col: int, value: Any):
    """
    安全填寫儲存格：若遇到 MergedCell (唯讀)，自動尋找並寫入該合併區塊最左上角的 Master Cell。
    """
    try:
        cell = ws.cell(row=row, column=col)
        if type(cell).__name__ == 'MergedCell':
            for rng in ws.merged_cells.ranges:
                if cell.coordinate in rng:
                    ws.cell(row=rng.min_row, column=rng.min_col, value=value)
                    return
        cell.value = value
    except Exception:
        pass


def parse_cad_with_screenshot(file_path: str) -> Dict[str, Any]:
    """
    讀取 CAD 檔案 (.step, .stp, .igs, .iges)，提取邊界尺寸，並以獨立 try-except 保護截圖生成。
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到檔案：{file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    valid_extensions = ('.step', '.stp', '.igs', '.iges')
    if ext not in valid_extensions:
        raise ValueError(f"不支援的格式 '{ext}'")

    # 1. 核心幾何長寬高解析 (方案 A: 完整支援 STEP 與 IGES/IGS)
    try:
        if ext in ('.step', '.stp'):
            model = cq.importers.importStep(file_path)
        elif ext in ('.igs', '.iges'):
            try:
                # 優先嘗試 CadQuery 標準形狀匯入
                model = cq.importers.importShape(file_path)
            except Exception:
                # 降級方案：調用底層 OpenCASCADE (OCP) 控制引擎讀取 IGES
                from OCP.IGESControl import IGESControl_Reader
                reader = IGESControl_Reader()
                reader.ReadFile(file_path)
                reader.TransferRoots()
                occ_shape = reader.Shape()
                model = cq.Workplane("XY").newObject([cq.Shape.cast(occ_shape)])

        bbox = model.val().BoundingBox()
        x_len = round(bbox.xlen, 2)
        y_len = round(bbox.ylen, 2)
        z_len = round(bbox.zlen, 2)
        dims_sorted: List[float] = sorted([x_len, y_len, z_len], reverse=True)
        
        length_val = dims_sorted[0]
        width_val = dims_sorted[1]
        height_val = dims_sorted[2]
        
        dimensions_str = f"{length_val:.2f}*{width_val:.2f}*{height_val:.2f}"
    except Exception as e:
        return {
            "status": "error",
            "file_name": os.path.basename(file_path),
            "error_message": f"CAD 幾何解析失敗: {str(e)}"
        }

    # 2. 獨立截圖流程
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
    載入範本檔，寫入解析資料並安全修復第 101 列合計公式。
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

    # 2. 原地寫入資料（由第 10 列起）
    start_row = 10
    for idx, item in enumerate(parsed_results):
        row_num = start_row + idx
        
        # A 欄: 序項
        ws.cell(row=row_num, column=1, value=idx + 1)
        
        # B 欄: 品名
        ws.cell(row=row_num, column=2, value=item.get("file_name", ""))
        
        # P 欄: 尺寸整合字串 (長*寬*高)
        ws.cell(row=row_num, column=16, value=item.get("dimensions_str", ""))
        
        # Q, R, S 欄: 長、寬、高拆解數值
        if "length" in item:
            ws.cell(row=row_num, column=17, value=item["length"])
        if "width" in item:
            ws.cell(row=row_num, column=18, value=item["width"])
        if "height" in item:
            ws.cell(row=row_num, column=19, value=item["height"])
            
        # T 欄: 單位 (mm)
        ws.cell(row=row_num, column=20, value=item.get("unit", "mm"))

        # 保留與設定明細金額計算公式 (數量 * 單價)
        ws.cell(row=row_num, column=7, value=f"=E{row_num}*F{row_num}")  # G欄: 原型金額
        ws.cell(row=row_num, column=10, value=f"=H{row_num}*I{row_num}") # J欄: 矽膠模具金額
        ws.cell(row=row_num, column=13, value=f"=K{row_num}*L{row_num}") # M欄: 注型金額

    # 3. 動態尋找「合計」列號（預設第 101 列）
    total_row = 101
    for r in range(10, ws.max_row + 1):
        cell_a = str(ws.cell(row=r, column=1).value or "").replace(" ", "")
        cell_b = str(ws.cell(row=r, column=2).value or "").replace(" ", "")
        cell_c = str(ws.cell(row=r, column=3).value or "").replace(" ", "")
        
        if "合計" in cell_a or "合計" in cell_b or "合計" in cell_c:
            total_row = r
            break

    # 4. 使用 set_cell_value_safe 安全寫入合計加總公式
    data_end_row = total_row - 1
    set_cell_value_safe(ws, total_row, 7, f"=SUM(G10:G{data_end_row})")    # G欄: 原型金額合計
    set_cell_value_safe(ws, total_row, 10, f"=SUM(J10:J{data_end_row})")   # J欄: 模具金額合計
    set_cell_value_safe(ws, total_row, 13, f"=SUM(M10:M{data_end_row})")   # M欄: 注型金額合計
    set_cell_value_safe(ws, total_row, 15, f"=G{total_row}+J{total_row}+M{total_row}") # O欄: 總金額合計

    wb.save(output_excel_path)
