"""
cad_parser.py - CAD 解析核心模組
支援讀取 .step / .stp / .igs / .iges 檔案並計算 Bounding Box 邊界外框尺寸。
"""

import os
from typing import Dict, Any, List
import cadquery as cq


def parse_cad_bounding_box(file_path: str) -> Dict[str, Any]:
    """
    讀取指定路徑的 CAD 檔案並回傳長寬高尺寸資料。

    :param file_path: CAD 檔案路徑
    :return: 包含原始尺寸與排序後胚料尺寸的字典
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到指定的檔案：{file_path}")

    # 驗證副檔名
    valid_extensions = ('.step', '.stp', '.igs', '.iges')
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in valid_extensions:
        raise ValueError(f"不支援的檔案格式 '{ext}'。僅支援：{', '.join(valid_extensions)}")

    try:
        # 載入 3D 模型
        model = cq.importers.importShape(file_path)

        # 取得 Bounding Box
        bbox = model.val().BoundingBox()

        # 提取原始 X, Y, Z 軸尺寸 (保留小數點後兩位，單位為 mm)
        x_len = round(bbox.xlen, 2)
        y_len = round(bbox.ylen, 2)
        z_len = round(bbox.zlen, 2)

        # 依製造業習慣排序：長 >= 寬 >= 高
        dims_sorted: List[float] = sorted([x_len, y_len, z_len], reverse=True)

        return {
            "status": "success",
            "file_name": os.path.basename(file_path),
            "file_path": os.path.abspath(file_path),
            "raw_dimensions": {
                "x_axis_mm": x_len,
                "y_axis_mm": y_len,
                "z_axis_mm": z_len,
            },
            "stock_dimensions": {
                "length_mm": dims_sorted[0],
                "width_mm": dims_sorted[1],
                "height_mm": dims_sorted[2],
            }
        }

    except Exception as e:
        return {
            "status": "error",
            "file_name": os.path.basename(file_path),
            "error_message": str(e)
        }