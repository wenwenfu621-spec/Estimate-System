"""
pages/2_📄_2D_Mylar_Analysis.py - 2D Mylar / DXF 材料辨識專屬頁面
Version: v2.8.0_20260821
Description: 專責 DXF 上傳、Mylar 厚度與材質設定、2D 輪廓解析與預覽。
             使用獨立的 2D Session State 與專屬 Reset 邏輯。
"""

import streamlit as st
import os
import tempfile
from datetime import datetime
from dxf_parser import parse_dxf_2d
from cad_parser import generate_excel_report, generate_word_report, is_valid_image


def cleanup_2d_temp():
    if "temp_files_2d" in st.session_state and st.session_state.temp_files_2d:
        for fpath in st.session_state.temp_files_2d:
            if fpath and os.path.exists(fpath):
                try:
                    os.remove(fpath)
                except Exception:
                    pass
        st.session_state.temp_files_2d = []


def reset_2d_session():
    cleanup_2d_temp()
    st.session_state.parsed_2d_results = None
    st.session_state.excel_2d_bytes = None
    st.session_state.word_2d_bytes = None
    st.session_state.uploader_2d_key += 1


st.set_page_config(page_title="2D Mylar 材料辨識", page_icon="📄", layout="centered")
st.title("📄 2D Mylar / 模切材料辨識")
st.write("上傳 `.dxf` 2D 模切圖檔（DWG 請先另存為 DXF），設定材料厚度與類型以計算外形尺寸與面積。")

if "uploader_2d_key" not in st.session_state:
    st.session_state.uploader_2d_key = 0
if "parsed_2d_results" not in st.session_state:
    st.session_state.parsed_2d_results = None
if "excel_2d_bytes" not in st.session_state:
    st.session_state.excel_2d_bytes = None
if "word_2d_bytes" not in st.session_state:
    st.session_state.word_2d_bytes = None
if "temp_files_2d" not in st.session_state:
    st.session_state.temp_files_2d = []

with st.expander("📝 客戶與報價表頭資訊填寫", expanded=True):
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        customer_input = st.text_input("客戶名稱", key=f"cust_2d_{st.session_state.uploader_2d_key}")
        phone_input = st.text_input("聯絡電話", key=f"phone_2d_{st.session_state.uploader_2d_key}")
    with col_c2:
        contact_input = st.text_input("聯絡人", key=f"contact_2d_{st.session_state.uploader_2d_key}")
        fax_input = st.text_input("傳真", key=f"fax_2d_{st.session_state.uploader_2d_key}")

with st.expander("🧪 2D 模切材料參數設定", expanded=True):
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        material_type = st.selectbox("材料類型 (Material Type)", ["Mylar", "PET", "PC Film", "Foam", "Sponge", "Rubber", "Gasket", "Tape", "Other"])
    with col_m2:
        thickness_option = st.selectbox("材料厚度 Thickness (mm)", [0.10, 0.125, 0.188, 0.20, 0.25, 0.30, 0.50, "自訂"])
        if thickness_option == "自訂":
            selected_thickness = st.number_input("輸入自訂厚度 (mm)", min_value=0.01, value=0.25, step=0.01)
        else:
            selected_thickness = float(thickness_option)

st.info("⚠️ **DWG 提示**：DWG 目前不直接解析，請先將 DWG 另存為 DXF 後再上傳。")

uploaded_2d_files = st.file_uploader(
    "上傳 2D DXF 圖檔 (可多選)", 
    type=["dxf"],
    accept_multiple_files=True,
    key=f"uploader_2d_{st.session_state.uploader_2d_key}"
)

if uploaded_2d_files:
    st.write(f"已選擇 **{len(uploaded_2d_files)}** 個 DXF 檔案。")

    if st.button("🎯 開始 2D 輪廓與面積辨識", type="primary", width="stretch"):
        cleanup_2d_temp()
        parsed_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, up_file in enumerate(uploaded_2d_files):
            fname = up_file.name
            status_text.text(f"正在處理 DXF 檔案 ({idx+1}/{len(uploaded_2d_files)}): {fname} ...")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
                tmp.write(up_file.getvalue())
                tmp_path = tmp.name
                st.session_state.temp_files_2d.append(tmp_path)

            res = parse_dxf_2d(tmp_path, user_thickness=selected_thickness)
            res["file_name"] = fname
            res["material_type"] = material_type
            parsed_results.append(res)

            if res.get("image_path"):
                st.session_state.temp_files_2d.append(res["image_path"])

            progress_bar.progress((idx + 1) / len(uploaded_2d_files))

        status_text.success("🎉 2D DXF 檔案全部辨識完成！")

        header_info = {"customer": customer_input, "contact": contact_input, "phone": phone_input, "fax": fax_input}

        excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        excel_path = excel_tmp.name
        excel_tmp.close()
        st.session_state.temp_files_2d.append(excel_path)

        generate_excel_report(parsed_results, excel_path, header_info=header_info)
        with open(excel_path, "rb") as f:
            excel_bytes = f.read()

        word_bytes = None
        try:
            word_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            word_path = word_tmp.name
            word_tmp.close()
            st.session_state.temp_files_2d.append(word_path)

            generate_word_report(parsed_results, word_path, header_info=header_info)
            with open(word_path, "rb") as f_w:
                word_bytes = f_w.read()
        except Exception as e_w:
            st.warning(f"⚠️ Word 產生警告: {str(e_w)}")

        st.session_state.parsed_2d_results = parsed_results
        st.session_state.excel_2d_bytes = excel_bytes
        st.session_state.word_2d_bytes = word_bytes

if st.session_state.parsed_2d_results:
    st.subheader("📋 2D Mylar 辨識結果預覽")
    for res in st.session_state.parsed_2d_results:
        if res.get("status") == "success":
            dims_esc = res['dimensions_str'].replace("*", "\\*")
            with st.expander(f"📄 {res['file_name']} - 【尺寸：{dims_esc} mm】"):
                col1, col2 = st.columns([2, 3])
                with col1:
                    st.write(f"**檔名**：{res['file_name']}")
                    st.write(f"**材料類型**：{res.get('material_type', 'Mylar')}")
                    st.write(f"**外形尺寸**：{dims_esc} mm")
                    st.write(f"**外形總面積**：{res.get('gross_area')} mm²")
                    if res.get('unit_warning'):
                        st.warning(res['unit_warning'])
                with col2:
                    if is_valid_image(res.get("image_path")):
                        st.image(res["image_path"], caption="2D 材料輪廓預覽", use_container_width=True)
                    else:
                        st.warning("⚠️ 預覽圖產生失敗")
        else:
            st.error(f"❌ {res['file_name']} 解析失敗：{res.get('error_message')}")

    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    today_str = datetime.now().strftime("%Y%m%d")
    with col1:
        st.download_button("📊 下載 Excel 報價單", st.session_state.excel_2d_bytes, f"2D_Report_{today_str}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    with col2:
        if st.session_state.word_2d_bytes:
            st.download_button("📝 下載圖文報告 (.docx)", st.session_state.word_2d_bytes, f"2D_Report_{today_str}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", width="stretch")
        else:
            st.button("📝 Word 不可用", disabled=True, width="stretch")
    with col3:
        st.button("🔄 重置 2D 頁面", on_click=reset_2d_session, width="stretch")
