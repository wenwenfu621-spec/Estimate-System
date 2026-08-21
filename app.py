"""
app.py - Streamlit 網頁介面程式
Version: v2.7.0_20260821
Description: 整合 3D (STEP/IGES) 與 2D (DXF) 雙軌路由管道，
             支援 Mylar 材料設定與 Thickness 輸入介面，
             底部按鈕維持等寬 (st.columns(3) + width="stretch")。
"""

import streamlit as st
import os
import tempfile
import base64
from datetime import datetime
from cad_parser import parse_cad_with_screenshot, generate_excel_report, generate_word_report, is_valid_image
from dxf_parser import parse_dxf_2d


def inject_custom_elements():
    """注入左下角版本號標籤與中央個人識別頭像徽章"""
    avatar_candidates = ["avatar.jpg", "avatar.jpeg", "avatar.png", "avatar.JPG", "avatar.PNG"]
    img_base64 = ""
    mime_type = "image/png"

    for af in avatar_candidates:
        if os.path.exists(af):
            with open(af, "rb") as img_f:
                img_base64 = base64.b64encode(img_f.read()).decode("utf-8")
                mime_type = "image/jpeg" if af.lower().endswith((".jpg", ".jpeg")) else "image/png"
            break

    avatar_html = f'<img src="data:{mime_type};base64,{img_base64}" style="width: 36px; height: 36px; border-radius: 50%; object-fit: cover; margin-right: 8px; border: 1.5px solid #ccc; background-color: #fff;">' if img_base64 else ""

    custom_css = f"""
    <style>
    .version-badge-left {{
        position: fixed; bottom: 16px; left: 16px;
        background-color: rgba(240, 242, 246, 0.9); padding: 4px 12px;
        border-radius: 12px; font-family: monospace; font-size: 0.8rem;
        color: #555555; border: 1px solid #d0d0d0; z-index: 999999; pointer-events: none;
    }}
    .custom-footer-max {{
        position: fixed; bottom: 16px; left: 50%; transform: translateX(-50%);
        display: flex; align-items: center; background-color: rgba(255, 255, 255, 0.95);
        padding: 4px 14px; border-radius: 20px; box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.15);
        z-index: 999999; pointer-events: none;
    }}
    .custom-footer-text {{
        font-family: 'Comic Sans MS', cursive, sans-serif; font-weight: bold;
        font-style: italic; font-size: 0.95rem; color: #333333; white-space: nowrap;
    }}
    .diagram-card-box {{
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 12px; text-align: center; margin-bottom: 12px;
    }}
    </style>
    <div class="version-badge-left">Version: v2.7.0_20260821</div>
    <div class="custom-footer-max">{avatar_html}<span class="custom-footer-text">Design by Max</span></div>
    """
    st.markdown(custom_css, unsafe_allow_html=True)


def cleanup_temp_files():
    if "temp_files_list" in st.session_state and st.session_state.temp_files_list:
        for fpath in st.session_state.temp_files_list:
            if fpath and os.path.exists(fpath):
                try:
                    os.remove(fpath)
                except Exception:
                    pass
        st.session_state.temp_files_list = []


def reset_session():
    cleanup_temp_files()
    st.session_state.parsed_results = None
    st.session_state.excel_bytes = None
    st.session_state.word_bytes = None
    st.session_state.export_ext = ".xlsx"
    st.session_state.mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.session_state.uploader_key_num += 1


st.set_page_config(page_title="CAD & Mylar 報價辨識工具 (v2.7.0)", page_icon="⚙️", layout="centered")
inject_custom_elements()

st.markdown("<h1 style='text-align: center; white-space: nowrap;'>⚙️ CAD 與 2D 模切自動報價系統 ⚙️</h1>", unsafe_allow_html=True)
st.write("支援 **3D 機構件 (STEP/IGES)** 與 **2D 模切材料 (DXF Mylar/PET)** 混合上傳辨識。")

if "uploader_key_num" not in st.session_state:
    st.session_state.uploader_key_num = 0
if "parsed_results" not in st.session_state:
    st.session_state.parsed_results = None
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "word_bytes" not in st.session_state:
    st.session_state.word_bytes = None
if "export_ext" not in st.session_state:
    st.session_state.export_ext = ".xlsx"
if "mime_type" not in st.session_state:
    st.session_state.mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
if "temp_files_list" not in st.session_state:
    st.session_state.temp_files_list = []

# 表頭與 Mylar 參數設定
with st.expander("📝 客戶與報價表頭資訊填寫", expanded=True):
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        customer_input = st.text_input("客戶名稱", key=f"cust_{st.session_state.uploader_key_num}")
        phone_input = st.text_input("聯絡電話", key=f"phone_{st.session_state.uploader_key_num}")
    with col_c2:
        contact_input = st.text_input("聯絡人", key=f"contact_{st.session_state.uploader_key_num}")
        fax_input = st.text_input("傳真", key=f"fax_{st.session_state.uploader_key_num}")

with st.expander("🧪 2D 模切材料預設參數 (針對 DXF)", expanded=True):
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        material_type = st.selectbox("材料類型 (2D Material)", ["Mylar", "PET", "PC Film", "Foam", "Sponge", "Rubber", "Gasket", "Tape", "Other"])
    with col_m2:
        thickness_option = st.selectbox("預設材料厚度 Thickness (mm)", [0.10, 0.125, 0.188, 0.20, 0.25, 0.30, 0.50, "自訂"])
        if thickness_option == "自訂":
            custom_thickness = st.number_input("輸入自訂厚度 (mm)", min_value=0.01, value=0.25, step=0.01)
            selected_thickness = custom_thickness
        else:
            selected_thickness = float(thickness_option)

uploaded_files = st.file_uploader(
    "上傳圖檔 (支援 STEP, STP, IGS, IGES, DXF；DWG 請先另存為 DXF)", 
    type=["step", "stp", "igs", "iges", "dxf", "dwg"],
    accept_multiple_files=True,
    key=f"file_uploader_{st.session_state.uploader_key_num}"
)

if uploaded_files:
    st.write(f"已選擇 **{len(uploaded_files)}** 個檔案。")

    if st.button("🎯 開始雙軌智慧辨識", type="primary"):
        cleanup_temp_files()
        parsed_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, uploaded_file in enumerate(uploaded_files):
            fname = uploaded_file.name
            ext = os.path.splitext(fname)[1].lower()
            status_text.text(f"正在辨識 ({idx+1}/{len(uploaded_files)}): {fname} ...")

            # DWG 攔截處理
            if ext == '.dwg':
                parsed_results.append({
                    "status": "error",
                    "file_name": fname,
                    "error_message": "DWG direct parsing is currently unsupported. 請將 DWG 另存為 DXF 後重新上傳。"
                })
                progress_bar.progress((idx + 1) / len(uploaded_files))
                continue

            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name
                st.session_state.temp_files_list.append(tmp_path)

            # 路由分流：3D vs 2D
            if ext in ('.dxf',):
                result = parse_dxf_2d(tmp_path, user_thickness=selected_thickness)
                result["material_type"] = material_type
            else:
                result = parse_cad_with_screenshot(tmp_path)

            result["file_name"] = fname
            parsed_results.append(result)

            if result.get("image_path"):
                st.session_state.temp_files_list.append(result["image_path"])

            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.success("🎉 所有檔案辨識與路由處理完成！")

        header_info = {"customer": customer_input, "contact": contact_input, "phone": phone_input, "fax": fax_input}

        has_xlsm_template = os.path.exists("template.xlsm") or os.path.exists("Template.xlsm")
        export_ext = ".xlsm" if has_xlsm_template else ".xlsx"
        mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12" if has_xlsm_template else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=export_ext)
        excel_path = excel_tmp.name
        excel_tmp.close()
        st.session_state.temp_files_list.append(excel_path)

        generate_excel_report(parsed_results, excel_path, header_info=header_info)
        with open(excel_path, "rb") as f:
            excel_bytes = f.read()

        word_bytes = None
        try:
            word_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            word_path = word_tmp.name
            word_tmp.close()
            st.session_state.temp_files_list.append(word_path)

            generate_word_report(parsed_results, word_path, header_info=header_info)
            with open(word_path, "rb") as f_w:
                word_bytes = f_w.read()
        except Exception as e_word:
            st.warning(f"⚠️ Word 報表產生警告：{str(e_word)}")

        st.session_state.parsed_results = parsed_results
        st.session_state.excel_bytes = excel_bytes
        st.session_state.word_bytes = word_bytes
        st.session_state.export_ext = export_ext
        st.session_state.mime_type = mime_type

if st.session_state.parsed_results:
    st.subheader("📋 雙軌辨識結果預覽")
    for res in st.session_state.parsed_results:
        if res.get("status") == "success":
            ftype = res.get("file_type", "3D")
            display_dims_escaped = res['dimensions_str'].replace("*", "\\*")
            
            with st.expander(f"📄 [{ftype}] {res['file_name']} - 【尺寸：{display_dims_escaped} mm】"):
                col_p1, col_p2 = st.columns([2, 3])
                with col_p1:
                    st.write(f"**檔名**：{res['file_name']}")
                    st.write(f"**類型**：{'2D 模切材料 (' + res.get('material_type', 'Mylar') + ')' if ftype == '2D' else '3D 機械機構件'}")
                    st.write(f"**尺寸**：{display_dims_escaped} mm")
                    if ftype == '2D':
                        st.write(f"**外形總面積**：{res.get('gross_area')} mm²")
                        if res.get('unit_warning'):
                            st.warning(res['unit_warning'])
                    else:
                        st.write(f"**演算法**：{res.get('used_mode', 'OBB')} 模式")
                with col_p2:
                    if is_valid_image(res.get("image_path")):
                        caption_text = "2D 材料輪廓預覽" if ftype == '2D' else "CAD 工程三視圖圖卡"
                        st.image(res["image_path"], caption=caption_text, use_container_width=True)
                    else:
                        st.warning("⚠️ 預覽圖產生失敗")
                        if res.get("image_error"):
                            st.code(f"Diagnose Error: {res['image_error']}", language="text")
        else:
            st.error(f"❌ {res['file_name']} 解析失敗：{res.get('error_message')}")

    st.markdown("---")
    
    col_dl1, col_dl2, col_rst = st.columns(3)
    today_date_str = datetime.now().strftime("%Y%m%d")
    
    with col_dl1:
        st.download_button(label="📊 下載 Excel 報價單", data=st.session_state.excel_bytes, file_name=f"Quotation_{today_date_str}{st.session_state.export_ext}", mime=st.session_state.mime_type, width="stretch")
    with col_dl2:
        if st.session_state.word_bytes:
            st.download_button(label="📝 下載圖文報價報告 (.docx)", data=st.session_state.word_bytes, file_name=f"Quotation_{today_date_str}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", width="stretch")
        else:
            st.button("📝 Word 報表不可用", disabled=True, width="stretch")
    with col_rst:
        st.button("🔄 重置 / 準備下一批", on_click=reset_session, width="stretch")
