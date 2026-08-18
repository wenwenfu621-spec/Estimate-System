"""
app.py - Streamlit 網頁介面程式
Version: v1.2.0_20260818
Description: 提供多檔 CAD 上傳、AI/CAD 按鈕觸發辨識、3D 圖形預覽與匯出 Excel 下載。
"""

import streamlit as st
import os
import tempfile
from cad_parser import parse_cad_with_screenshot, generate_excel_report

st.set_page_config(page_title="CAD 報價辨識工具", page_icon="⚙️", layout="centered")

st.title("⚙️ CAD 自動報價與尺寸辨識工具")
st.write("上傳 `.step` 或 `.igs` 3D 模型檔，點選下方按鈕自動辨識尺寸與擷取 3D 視角圖像。")

# 多檔案上傳器
uploaded_files = st.file_uploader(
    "上傳 CAD 圖檔 (可多選)", 
    type=["step", "stp", "igs", "iges"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"已選擇 **{len(uploaded_files)}** 個檔案。")

    # 1. AI / CAD 觸發辨識按鈕 (參考專案 UI 樣式)
    if st.button("🎯 辨識 CAD 單據與幾何內容", type="primary"):
        parsed_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"正在辨識 ({idx+1}/{len(uploaded_files)}): {uploaded_file.name} ...")
            
            ext = os.path.splitext(uploaded_file.name)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name

            # 執行解析與截圖
            result = parse_cad_with_screenshot(tmp_path)
            result["file_name"] = uploaded_file.name
            parsed_results.append(result)

            if os.path.exists(tmp_path):
                os.remove(tmp_path)

            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.success("🎉 所有檔案辨識完成！")

        # 2. 顯示辨識結果清單
        st.subheader("📋 辨識結果預覽")
        for res in parsed_results:
            if res.get("status") == "success":
                with st.expander(f"📄 {res['file_name']} - 【{res['dimensions_str']} mm】"):
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        if os.path.exists(res.get("image_path", "")):
                            st.image(res["image_path"], caption="3D 內定視角縮圖", width=150)
                    with col2:
                        st.write(f"**檔名**：{res['file_name']}")
                        st.write(f"**長寬高**：{res['dimensions_str']}")
                        st.write(f"**單位**：{res['unit']}")
            else:
                st.error(f"❌ {res['file_name']} 解析失敗：{res.get('error_message')}")

        # 3. 匯出 Excel
        excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        excel_path = excel_tmp.name
        excel_tmp.close()

        generate_excel_report(parsed_results, excel_path)

        with open(excel_path, "rb") as f:
            excel_bytes = f.read()

        st.markdown("---")
        # 下載 Excel 按鈕
        st.download_button(
            label="📊 下載 CAD 辨識結果 Excel 報表 (.xlsx)",
            data=excel_bytes,
            file_name="CAD_Quotation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 清理暫存圖片與 Excel
        if os.path.exists(excel_path):
            os.remove(excel_path)
        for res in parsed_results:
            img_p = res.get("image_path")
            if img_p and os.path.exists(img_p):
                os.remove(img_p)
