"""
pages/1_3D_CAD.py - 3D CAD 尺寸辨識專屬頁面
Version: v2.8.2_20260821
Description: 專責 STEP/IGES 上傳、OBB/AABB 計算與 3D 工程三視圖預覽。
             使用獨立的 3D Session State、專屬 Reset 邏輯與返回首頁導航。
"""

import streamlit as st
import os
import tempfile
from datetime import datetime
from cad_parser import parse_cad_with_screenshot, generate_excel_report, generate_word_report, is_valid_image


def cleanup_3d_temp():
    if "temp_files_3d" in st.session_state and st.session_state.temp_files_3d:
        for fpath in st.session_state.temp_files_3d:
            if fpath and os.path.exists(fpath):
                try:
                    os.remove(fpath)
                except Exception:
                    pass
        st.session_state.temp_files_3d = []


def reset_3d_session():
    cleanup_3d_temp()
    st.session_state.parsed_3d_results = None
    st.session_state.excel_3d_bytes = None
    st.session_state.word_3d_bytes = None
    st.session_state.uploader_3d_key += 1


st.set_page_config(page_title="3D CAD 尺寸辨識", page_icon="📦", layout="centered")

if st.button("← 返回功能首頁"):
    st.switch_page("app.py")

st.title("📦 3D 機構件尺寸與報價辨識")
st.write("上傳 `.step`, `.stp`, `.igs`, `.iges` 3D 模型檔，自動計算 OBB/AABB 包容盒與第三角法三視圖。")

if "uploader_3d_key" not in st.session_state:
    st.session_state.uploader_3d_key = 0
if "parsed_3d_results" not in st.session_state:
    st.session_state.parsed_3d_results = None
if "excel_3d_bytes" not in st.session_state:
    st.session_state.excel_3d_bytes = None
if "word_3d_bytes" not in st.session_state:
    st.session_state.word_3d_bytes = None
if "temp_files_3d" not in st.session_state:
    st.session_state.temp_files_3d = []

with st.expander("📝 客戶與報價表頭資訊填寫", expanded=True):
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        customer_input = st.text_input("客戶名稱", key=f"cust_3d_{st.session_state.uploader_3d_key}")
        phone_input = st.text_input("聯絡電話", key=f"phone_3d_{st.session_state.uploader_3d_key}")
    with col_c2:
        contact_input = st.text_input("聯絡人", key=f"contact_3d_{st.session_state.uploader_3d_key}")
        fax_input = st.text_input("傳真", key=f"fax_3d_{st.session_state.uploader_3d_key}")

uploaded_3d_files = st.file_uploader(
    "上傳 3D CAD 圖檔 (可多選)", 
    type=["step", "stp", "igs", "iges"],
    accept_multiple_files=True,
    key=f"uploader_3d_{st.session_state.uploader_3d_key}"
)

if uploaded_3d_files:
    st.write(f"已選擇 **{len(uploaded_3d_files)}** 個 3D 檔案。")

    if st.button("🎯 開始 3D 幾何與工程視圖辨識", type="primary", width="stretch"):
        cleanup_3d_temp()
        parsed_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, up_file in enumerate(uploaded_3d_files):
            fname = up_file.name
            ext = os.path.splitext(fname)[1]
            status_text.text(f"正在處理 3D 檔案 ({idx+1}/{len(uploaded_3d_files)}): {fname} ...")

            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                tmp.write(up_file.getvalue())
                tmp_path = tmp.name
                st.session_state.temp_files_3d.append(tmp_path)

            res = parse_cad_with_screenshot(tmp_path)
            res["file_name"] = fname
            parsed_results.append(res)

            if res.get("image_path"):
                st.session_state.temp_files_3d.append(res["image_path"])

            progress_bar.progress((idx + 1) / len(uploaded_3d_files))

        status_text.success("🎉 3D 檔案全部辨識完成！")

        header_info = {"customer": customer_input, "contact": contact_input, "phone": phone_input, "fax": fax_input}

        excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        excel_path = excel_tmp.name
        excel_tmp.close()
        st.session_state.temp_files_3d.append(excel_path)

        generate_excel_report(parsed_results, excel_path, header_info=header_info)
        with open(excel_path, "rb") as f:
            excel_bytes = f.read()

        word_bytes = None
        try:
            word_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            word_path = word_tmp.name
            word_tmp.close()
            st.session_state.temp_files_3d.append(word_path)

            generate_word_report(parsed_results, word_path, header_info=header_info)
            with open(word_path, "rb") as f_w:
                word_bytes = f_w.read()
        except Exception as e_w:
            st.warning(f"⚠️ Word 產生警告: {str(e_w)}")

        st.session_state.parsed_3d_results = parsed_results
        st.session_state.excel_3d_bytes = excel_bytes
        st.session_state.word_3d_bytes = word_bytes

if st.session_state.parsed_3d_results:
    st.subheader("📋 3D 辨識結果預覽")
    for res in st.session_state.parsed_3d_results:
        if res.get("status") == "success":
            dims_esc = res['dimensions_str'].replace("*", "\\*")
            mode = res.get("used_mode", "OBB")
            with st.expander(f"📄 {res['file_name']} - 【尺寸：{dims_esc} mm ({mode})】"):
                col1, col2 = st.columns([2, 3])
                with col1:
                    st.write(f"**檔名**：{res['file_name']}")
                    st.write(f"**尺寸**：{dims_esc} mm ({mode})")
                    st.write(f"**演算法**：{mode} 模式")
                    st.write(f"**拆解數值**：長 {res.get('length')} / 寬 {res.get('width')} / 高 {res.get('height')}")
                with col2:
                    if is_valid_image(res.get("image_path")):
                        st.image(res["image_path"], caption="CAD 工程三視圖圖卡 (第三角法)", use_container_width=True)
                    else:
                        st.warning("⚠️ 預覽圖產生失敗")
        else:
            st.error(f"❌ {res['file_name']} 解析失敗：{res.get('error_message')}")

    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    today_str = datetime.now().strftime("%Y%m%d")
    with col1:
        st.download_button("📊 下載 Excel 報價單", st.session_state.excel_3d_bytes, f"3D_Report_{today_str}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    with col2:
        if st.session_state.word_3d_bytes:
            st.download_button("📝 下載圖文報告 (.docx)", st.session_state.word_3d_bytes, f"3D_Report_{today_str}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", width="stretch")
        else:
            st.button("📝 Word 不可用", disabled=True, width="stretch")
    with col3:
        st.button("🔄 重置 3D 頁面", on_click=reset_3d_session, width="stretch")
