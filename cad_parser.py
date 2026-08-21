"""
cad_parser.py - CAD 解析與報表產出模組
Version: v2.3.1_20260821
Description: 支援載入 template.xlsm / template.xlsx 範本檔。
             修正自我循環引用 ImportError 錯誤。
             採用 OpenCASCADE 原生 Bnd_OBB API 計算精確最小包容盒素材尺寸，
             尺寸採 math.ceil 無條件進位至個位數整數。
             使用 safe_str + set_cell_value_safe 安全寫入 B4~B7 表頭資訊，
             B/P 欄自動調整欄寬，全數儲存格統一指定為「標楷體」。
"""

import os
import math
import tempfile
from datetime import datetime
from typing import Dict, Any, List, Optional
import cadquery as cq
from PIL import Image
import openpyxl
from openpyxl.styles import Font

# 引入 OpenCASCADE 原生 Bnd_OBB 模組以進行精確幾何最小包容盒計算
try:
    from OCP.Bnd import Bnd_OBB
    from OCP.BRepBndLib import BRepBndLib
    HAS_OCP_OBB = True
except ImportError:
    HAS_OCP_OBB = False


def safe_str(val: Any) -> str:
    """
    安全轉為字串並去除首尾空白，防止 None 或特殊物件觸發 AttributeError
    """
    if val is None:
        return ""
    try:
        return str(val).strip()
    except Exception:
        return ""


def set_cell_value_safe(ws, row: int, col: int, value: Any, font: Optional[Font] = None):
    """
    安全填寫儲存格：若遇到 MergedCell (唯讀)，自動尋找並寫入該合併區塊最左上角的 Master Cell，
    並可同步套用指定字體。
    """
    try:
        cell = ws.cell(row=row, column=col)
        if type(cell).__name__ == 'MergedCell':
            for rng in ws.merged_cells.ranges:
                if cell.coordinate in rng:
                    master_cell = ws.cell(row=rng.min_row, column=rng.min_col)
                    master_cell.value = value
                    if font:
                        master_cell.font = font
                    return
        cell.value = value
        if font:
            cell.font = font
    except Exception:
        pass


def calculate_obb_dimensions(model: cq.Workplane) -> List[float]:
    """
    使用 OpenCASCADE 原生 Bnd_OBB 精確計算 3D 實體的最小素材包容盒 (OBB)
    """
    if HAS_OCP_OBB:
        try:
            occ_shape = model.val().wrapped
            obb = Bnd_OBB()
            # 建立精確之 Optimal Bounding Box
            BRepBndLib.AddOBB_s(occ_shape, obb, True, True, True)
            
            # XSize, YSize, ZSize 即為 OBB 之全尺寸 (長/寬/高)
            dims = [obb.XSize(), obb.YSize(), obb.ZSize()]
            if all(d > 0 for d in dims):
                return dims
        except Exception:
            pass

    # 若計算失敗或缺乏 OCP 模組，備援退回基本 BoundingBox
    bbox = model.val().BoundingBox()
    return [bbox.xlen, bbox.ylen, bbox.zlen]


def parse_cad_with_screenshot(file_path: str, mode: str = "OBB") -> Dict[str, Any]:
    """
    讀取 CAD 檔案，依指定模式 (OBB 或 AABB) 提取邊界尺寸，並無條件進位至個位數整數 (math.ceil)。
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到檔案：{file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    valid_extensions = ('.step', '.stp', '.igs', '.iges')
    if ext not in valid_extensions:
        raise ValueError(f"不支援的格式 '{ext}'")

    # 1. 幾何長寬高解析與整數進位處理
    try:
        if ext in ('.step', '.stp'):
            model = cq.importers.importStep(file_path)
        elif ext in ('.igs', '.iges'):
            try:
                model = cq.importers.importShape(file_path)
            except Exception:
                from OCP.IGESControl import IGESControl_Reader
                reader = IGESControl_Reader()
                reader.ReadFile(file_path)
                reader.TransferRoots()
                occ_shape = reader.Shape()
                model = cq.Workplane("XY").newObject([cq.Shape.cast(occ_shape)])

        if mode == "OBB":
            raw_dims = calculate_obb_dimensions(model)
        else:
            bbox = model.val().BoundingBox()
            raw_dims = [bbox.xlen, bbox.ylen, bbox.zlen]

        # 無條件進位至個位數整數
        x_len = math.ceil(raw_dims[0])
        y_len = math.ceil(raw_dims[1])
        z_len = math.ceil(raw_dims[2])

        dims_sorted: List[int] = sorted([x_len, y_len, z_len], reverse=True)

        length_val = dims_sorted[0]
        width_val = dims_sorted[1]
        height_val = dims_sorted[2]

        dimensions_str = f"{length_val}*{width_val}*{height_val}"
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


def generate_excel_report(
    parsed_results: List[Dict[str, Any]], 
    output_excel_path: str,
    header_info: Optional[Dict[str, Any]] = None
):
    """
    載入範本檔，寫入客戶表頭資訊與解析數據，全數指定字體為標楷體，並自動調整 B/P 欄寬。
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

    # 統一標楷體 Font 物件
    kai_font_regular = Font(name="標楷體", size=11, bold=False)
    kai_font_bold = Font(name="標楷體", size=11, bold=True)

    # 1. 使用 safe_str + set_cell_value_safe 安全寫入 B4~B7 表頭客戶資訊
    if header_info and isinstance(header_info, dict):
        customer_val = safe_str(header_info.get("customer"))
        contact_val = safe_str(header_info.get("contact"))
        phone_val = safe_str(header_info.get("phone"))
        fax_val = safe_str(header_info.get("fax"))

        if customer_val:
            set_cell_value_safe(ws, 4, 2, f"客戶名稱 : {customer_val}", font=kai_font_regular)
            
        if contact_val:
            set_cell_value_safe(ws, 5, 2, f"聯 絡 人 : {contact_val}", font=kai_font_regular)

        if phone_val:
            set_cell_value_safe(ws, 6, 2, f"聯絡電話 : {phone_val}", font=kai_font_regular)

        if fax_val:
            set_cell_value_safe(ws, 7, 2, f"傳    真 : {fax_val}", font=kai_font_regular)

    # 2. 填入當天報價日期於 O7 欄位
    today_str = datetime.now().strftime("%Y/%m/%d")
    set_cell_value_safe(ws, 7, 15, today_str, font=kai_font_regular)

    # 3. 寫入資料（由第 10 列起）並記錄最大文字長度以調整欄寬
    start_row = 10
    max_b_len = 12  # 品名欄最小字元計算基準
    max_p_len = 16  # 尺寸欄最小字元計算基準

    for idx, item in enumerate(parsed_results):
        row_num = start_row + idx
        
        # A 欄: 序項
        cell_a = ws.cell(row=row_num, column=1, value=idx + 1)
        cell_a.font = kai_font_regular
        
        # B 欄: 品名
        file_name = item.get("file_name", "")
        cell_b = ws.cell(row=row_num, column=2, value=file_name)
        cell_b.font = kai_font_regular
        max_b_len = max(max_b_len, len(str(file_name)))
        
        # P 欄: 尺寸整合字串 (長*寬*高)
        dims_str = item.get("dimensions_str", "")
        cell_p = ws.cell(row=row_num, column=16, value=dims_str)
        cell_p.font = kai_font_bold
        max_p_len = max(max_p_len, len(str(dims_str)))
        
        # Q, R, S 欄: 長、寬、高拆解數值 (統一標楷體粗體)
        if "length" in item:
            cell_q = ws.cell(row=row_num, column=17, value=item["length"])
            cell_q.font = kai_font_bold
        if "width" in item:
            cell_r = ws.cell(row=row_num, column=18, value=item["width"])
            cell_r.font = kai_font_bold
        if "height" in item:
            cell_s = ws.cell(row=row_num, column=19, value=item["height"])
            cell_s.font = kai_font_bold
            
        # T 欄: 單位 (mm)
        cell_t = ws.cell(row=row_num, column=20, value=item.get("unit", "mm"))
        cell_t.font = kai_font_regular

        # 明細金額計算公式
        c_g = ws.cell(row=row_num, column=7, value=f"=E{row_num}*F{row_num}")
        c_j = ws.cell(row=row_num, column=10, value=f"=H{row_num}*I{row_num}")
        c_m = ws.cell(row=row_num, column=13, value=f"=K{row_num}*L{row_num}")
        c_g.font = kai_font_regular
        c_j.font = kai_font_regular
        c_m.font = kai_font_regular

    # 4. 動態欄寬自動保護
    ws.column_dimensions['B'].width = max(ws.column_dimensions['B'].width or 0, max_b_len + 6)
    ws.column_dimensions['P'].width = max(ws.column_dimensions['P'].width or 0, max_p_len + 6)

    # 5. 動態尋找「合計」列號（預設第 101 列）
    total_row = 101
    for r in range(10, ws.max_row + 1):
        cell_a_val = str(ws.cell(row=r, column=1).value or "").replace(" ", "")
        cell_b_val = str(ws.cell(row=r, column=2).value or "").replace(" ", "")
        cell_c_val = str(ws.cell(row=r, column=3).value or "").replace(" ", "")
        
        if "合計" in cell_a_val or "合計" in cell_b_val or "合計" in cell_c_val:
            total_row = r
            break

    # 6. 使用 set_cell_value_safe 安全寫入合計加總公式
    data_end_row = total_row - 1
    set_cell_value_safe(ws, total_row, 7, f"=SUM(G10:G{data_end_row})", font=kai_font_bold)
    set_cell_value_safe(ws, total_row, 10, f"=SUM(J10:J{data_end_row})", font=kai_font_bold)
    set_cell_value_safe(ws, total_row, 13, f"=SUM(M10:M{data_end_row})", font=kai_font_bold)
    set_cell_value_safe(ws, total_row, 15, f"=G{total_row}+J{total_row}+M{total_row}", font=kai_font_bold)

    wb.save(output_excel_path)
