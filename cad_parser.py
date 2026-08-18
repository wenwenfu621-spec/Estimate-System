"""
cad_parser.py - CAD 解析與報表產出模組
Version: v1.2.0_20260818
Description: 支援 .step/.igs 解析長寬高、等角視角截圖繪製，並產出包含圖片的 Excel 報表。
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
    讀取 CAD 檔案，提取邊界尺寸並生成 3D 視角截圖。
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到檔案：{file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    valid_extensions = ('.step', '.stp', '.igs', '.iges')
    if ext not in valid_extensions:
        raise ValueError(f"不支援的格式 '{ext}'")

    try:
        # 1. 匯入 CAD 模型
        if ext in ('.step', '.stp'):
            model = cq.importers.importStep(file_path)
        else:
            model = cq.importers.importShape(cq.importers.ImportTypes.IGES, file_path)

        # 2. 取得 Bounding Box 尺寸 (單位: mm)
        bbox = model.val().BoundingBox()
        x_len = round(bbox.xlen, 2)
        y_len = round(bbox.ylen, 2)
        z_len = round(bbox.zlen, 2)
        dims_sorted: List[float] = sorted([x_len, y_len, z_len], reverse=True)

        # 3. 生成 3D 視角截圖 (等角視角 Isometric)
        img_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
        img_path = img_tmp.name
        img_tmp.close()

        # 匯出為 SVG 後轉成 PNG 縮圖
        svg_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.svg')
        svg_path = svg_tmp.name
        svg_tmp.close()

        cq.exporters.export(model, svg_path, opt={"projectionDir": (1, 1, 1), "showHidden": False})
        
        # 轉換為標準尺寸圖片 (寬 400x300 px)
        with Image.open(svg_path) as img:
            img_resized = img.resize((400, 300))
            img_resized.save(img_path, format="PNG")

        if os.path.exists(svg_path):
            os.remove(svg_path)

        return {
            "status": "success",
            "file_name": os.path.basename(file_path),
            "dimensions_str": f"{dims_sorted[0]:.2f} x {dims_sorted[1]:.2f} x {dims_sorted[2]:.2f}",
            "unit": "mm",
            "image_path": img_path
        }

    except Exception as e:
        return {
            "status": "error",
            "file_name": os.path.basename(file_path),
            "error_message": str(e)
        }


def generate_excel_report(parsed_results: List[Dict[str, Any]], output_excel_path: str):
    """
    將解析結果寫入 Excel，包含：項次 / 檔名 / 3D視角截圖 / 長寬高 / 單位
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CAD報價尺寸表"

    # 表頭設定
    headers = ["項次", "檔名", "3D 視角截圖", "長寬高", "單位"]
    ws.append(headers)

    # 樣式設定
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

    # 設定欄寬
    ws.column_dimensions['A'].width = 8   # 項次
    ws.column_dimensions['B'].width = 25  # 檔名
    ws.column_dimensions['C'].width = 16  # 3D 視角截圖 (縮圖預留)
    ws.column_dimensions['D'].width = 22  # 長寬高
    ws.column_dimensions['E'].width = 10  # 單位

    # 美化表頭
    for col_num in range(1, 6):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    # 寫入資料列
    for idx, item in enumerate(parsed_results, start=1):
        row_num = idx + 1
        
        # 設定固定列高 (70pt 容納圖片)
        ws.row_dimensions[row_num].height = 70

        ws.cell(row=row_num, column=1, value=idx)
        ws.cell(row=row_num, column=2, value=item.get("file_name", ""))
        ws.cell(row=row_num, column=4, value=item.get("dimensions_str", ""))
        ws.cell(row=row_num, column=5, value=item.get("unit", "mm"))

        # 套用儲存格格式
        for col_num in range(1, 6):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.alignment = center_align
            cell.border = thin_border

        # 插入 3D 視角截圖 (控制尺寸：寬 90px, 高 60px)
        img_path = item.get("image_path")
        if img_path and os.path.exists(img_path):
            img = OpenpyxlImage(img_path)
            img.width = 90
            img.height = 60
            # 放置於 C 欄對應格
            cell_address = f"C{row_num}"
            ws.add_image(img, cell_address)

    wb.save(output_excel_path)
