"""
cad_parser.py - CAD 解析與報表產出模組
Version: v1.4.0_20260818
Description: 方案 B 容錯解耦版。優先保障長寬高幾何尺寸解析與 Excel 匯出。
             截圖功能獨立防護，不因渲染例外而阻斷核心流程。
"""

import os
import tempfile
from typing import Dict, Any, List
import cadquery as cq
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


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

    # 1. 核心幾何長寬高解析（最優先保護）
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
        dimensions_str = f"{dims_sorted[0]:.2f} x {dims_sorted[1]:.2f} x {dims_sorted[2]:.2f}"
    except Exception as e:
        return {
            "status": "error",
            "file_name": os.path.basename(file_path),
            "error_message": f"CAD 幾何解析失敗: {str(e)}"
        }

    # 2. 獨立截圖流程（方案 B 容錯防護）
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
    將解析結果寫入 Excel，包含：項次 / 檔名 / 3D視角截圖 / 長寬高 / 單位
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CAD報價尺寸表"

    headers = ["項次", "檔名", "3D 視角截圖", "長寬高", "單位"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="微軟正黑體", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ws.column_dimensions['A'].width = 8   # 項次
    ws.column_dimensions['B'].width = 25  # 檔名
    ws.column_dimensions['C'].width = 16  # 3D 視角截圖
    ws.column_dimensions['D'].width = 22  # 長寬高
    ws.column_dimensions['E'].width = 10  # 單位

    for col_num in range(1, 6):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    for idx, item in enumerate(parsed_results, start=1):
        row_num = idx + 1
        ws.row_dimensions[row_num].height = 70

        ws.cell(row=row_num, column=1, value=idx)
        ws.cell(row=row_num, column=2, value=item.get("file_name", ""))
        ws.cell(row=row_num, column=4, value=item.get("dimensions_str", ""))
        ws.cell(row=row_num, column=5, value=item.get("unit", "mm"))

        for col_num in range(1, 6):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.alignment = center_align
            cell.border = thin_border

        img_path = item.get("image_path")
        if img_path and os.path.exists(img_path):
            try:
                img = OpenpyxlImage(img_path)
                img.width = 90
                img.height = 60
                ws.add_image(img, f"C{row_num}")
            except Exception:
                ws.cell(row=row_num, column=3, value="[無預覽圖]")
        else:
            ws.cell(row=row_num, column=3, value="[無預覽圖]")

    wb.save(output_excel_path)
