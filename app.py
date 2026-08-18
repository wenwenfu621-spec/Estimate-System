import streamlit as st
import os
import tempfile
from cad_parser import parse_cad_bounding_box

# 頁面標題與設定
st.set_page_config(page_title="CAD 尺寸解析工具", page_icon="📦", layout="centered")
st.title("📦 CAD 邊界外框尺寸解析工具")
st.write("請上傳 `.step` 或 `.igs` 3D 模型檔，系統將自動計算邊界外框與建議胚料尺寸。")

# 檔案上傳元件
uploaded_file = st.file_uploader(
    "拖曳或點擊上傳 CAD 檔案", 
    type=["step", "stp", "igs", "iges"]
)

if uploaded_file is not None:
    # 建立臨時檔案以供 cad_parser 讀取
    ext = os.path.splitext(uploaded_file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name

    st.info(f"正在解析檔案：**{uploaded_file.name}** ...")

    # 呼叫解析邏輯
    result = parse_cad_bounding_box(tmp_path)

    # 清理臨時檔案
    if os.path.exists(tmp_path):
        os.remove(tmp_path)

    # 顯示解析結果
    if result.get("status") == "success":
        st.success("✅ 解析成功！")
        stock = result["stock_dimensions"]
        raw = result["raw_dimensions"]

        st.subheader("📏 估算胚料尺寸 (長 ≥ 寬 ≥ 高)")
        col1, col2, col3 = st.columns(3)
        col1.metric("長度 (Length)", f"{stock['length_mm']} mm")
        col2.metric("寬度 (Width)", f"{stock['width_mm']} mm")
        col3.metric("高度 (Height)", f"{stock['height_mm']} mm")

        with st.expander("檢視原始座標軸尺寸與 JSON 數據"):
            st.write(f"原始尺寸：X = {raw['x_axis_mm']} mm, Y = {raw['y_axis_mm']} mm, Z = {raw['z_axis_mm']} mm")
            st.json(result)
    else:
        st.error(f"❌ 解析失敗：{result.get('error_message')}")