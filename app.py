"""
app.py - Streamlit 網頁介面程式
Version: v1.8.5_20260819
Description: 提供多檔 CAD 上傳與尺寸辨識，自動套寫至 Excel 範本。
             優先安全載入 UI 元素（左下角版本號、頁尾頭像徽章），
             徹底解決崩潰問題，支援一鍵重置與日期檔名。
"""

import streamlit as st
import os
import tempfile
import base64
from datetime import datetime
from cad_parser import parse_cad_with_screenshot, generate_excel_report


def inject_custom_elements():
    """注入左下角版本號標籤與中央個人識別頭像徽章"""
    avatar_candidates = [
        "avatar.jpg", "avatar.jpeg", "avatar.png", "avatar.JPG", "avatar.PNG",
        "Avatar.jpg", "Avatar.jpeg", "Avatar.png", "Avatar.JPG", "Avatar.PNG"
    ]
    img_base64 = ""
    mime_type = "image/png"

    for af in avatar_candidates:
        if os.path.exists(af):
            with open(af, "rb") as img_f:
                img_base64 = base64.b64encode(img_f.read()).decode("utf-8")
                if af.lower().endswith((".jpg", ".jpeg")):
                    mime_type = "image/jpeg"
                else:
                    mime_type = "image/png"
            break

    avatar_html = f'<img src="data:{mime_type};base64,{img_base64}" style="width: 36px; height: 36px; border-radius: 50%; object-fit: cover; margin-right: 8px; border: 1.5px solid #ccc; background-color: #fff;">' if img_base64 else ""

    custom_css = f"""
    <style>
    /* 左下角懸浮版本號 */
    .version-badge-left {{
        position: fixed;
        bottom: 16px;
        left: 16px;
        background-color: rgba(240, 242, 246, 0.9);
        padding: 4px 12px;
        border-radius: 12px;
        font-family: monospace;
        font-size: 0.8rem;
        color: #555555;
        border: 1px solid #d0d0d0;
        z-index: 999999;
        pointer-events: none;
    }}
    
    /* 底部正中央個人頭像徽章 */
    .custom-footer-max {{
        position: fixed;
        bottom: 16px;
        left: 50%;
        transform: translateX(-50%);
        display: flex;
        align-items: center;
        background-color: rgba(255, 255, 255, 0.95);
        padding: 4px 14px;
        border-radius: 20px;
        box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.15);
        z-index: 999999;
        pointer-events: none;
    }}
    .custom-footer-text {{
        font-family: 'Comic Sans MS', cursive, sans-serif;
        font-weight: bold;
        font-style: italic;
        font-size: 0.95rem;
        color: #333333;
        white-space: nowrap;
    }}
    </style>
    
    <div class="version-badge-left">Version: v1.8.5_20260819</div>
    
    <div class="custom-footer-max">
        {avatar_html}
        <span class="custom-footer-text">Design by Max</span>
    </div>
    """
    st.markdown(custom_css, unsafe_allow_html=True)


def cleanup_temp_files():
    """清理伺服器上記錄的當次暫存檔案"""
    if "temp_files_list" in st.session_state and st.session_state.temp_files_list:
        for fpath in st.session_state.temp_files_list:
            if fpath and os.path.exists(fpath):
                try:
                    os.remove(fpath)
                except Exception:
                    pass
        st.session_state.temp_files_list = []


def reset_session():
    """點擊重置按鈕時觸發：清空暫存檔、Session 狀態，並變更 uploader_key 強制清空上傳元件"""
    cleanup_temp_files()
    st.session_state.parsed_results = None
    st.session_state.excel_bytes = None
    st.session_state.export_ext = ".xlsx"
    st.session_state.mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    
    st.session_state.uploader_key_num += 1


st.set_page_config(page_title="CAD 報價辨識工具 (v1.8.5)", page_icon="⚙️", layout="centered")

# 優先注入左下角版本別與頁尾徽章，確保穩定顯示
inject_custom_elements()

st.title("CAD 自動報價與尺寸辨識工具")
st.write("上傳 `.step` 或 `.igs` 3D 模型檔，點選下方按鈕自動辨識尺寸並套寫至 Excel 報價單。")

if "uploader_key_num" not in st.session_state:
    st.session_state.uploader_key_num = 0
if "parsed_results" not in st.session_state:
    st.session_state.parsed_results = None
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "export_ext" not in st.session_state:
    st.session_state.export_ext = ".xlsx"
if "mime_type" not in st.session_state:
    st.session_state.mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
if "temp_files_list" not in st.session_state:
    st.session_state.temp_files_list = []

uploaded_files = st.file_uploader(
    "上傳 CAD 圖檔 (可多選)", 
    type=["step", "stp", "igs", "iges"],
    accept_multiple_files=True,
    key=f"file_uploader_{st.session_state.uploader_key_num}"
)

if uploaded_files:
    st.write(f"已選擇 **{len(uploaded_files)}** 個檔案。")

    if st.button("🎯 辨識 CAD 單據與幾何內容", type="primary"):
        cleanup_temp_files()

        parsed_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"正在辨識 ({idx+1}/{len(uploaded_files)}): {uploaded_file.name} ...")
            
            ext = os.path.splitext(uploaded_file.name)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name
                st.session_state.temp_files_list.append(tmp_path)

            result = parse_cad_with_screenshot(tmp_path)
            result["file_name"] = uploaded_file.name
            parsed_results.append(result)

            if result.get("image_path"):
                st.session_state.temp_files_list.append(result["image_path"])

            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.success("🎉 所有檔案辨識完成！已自動套用範本檔填入數據。")

        has_xlsm_template = os.path.exists("template.xlsm") or os.path.exists("Template.xlsm")
        export_ext = ".xlsm" if has_xlsm_template else ".xlsx"
        mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12" if has_xlsm_template else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=export_ext)
        excel_path = excel_tmp.name
        excel_tmp.close()
        st.session_state.temp_files_list.append(excel_path)

        generate_excel_report(parsed_results, excel_path)

        with open(excel_path, "rb") as f:
            excel_bytes = f.read()

        st.session_state.parsed_results = parsed_results
        st.session_state.excel_bytes = excel_bytes
        st.session_state.export_ext = export_ext
        st.session_state.mime_type = mime_type

if st.session_state.parsed_results:
    st.subheader("📋 辨識結果預覽")
    for res in st.session_state.parsed_results:
        if res.get("status") == "success":
            with st.expander(f"📄 {res['file_name']} - 【尺寸：{res['dimensions_str']} mm】"):
                st.write(f"**品名 (檔名)**：{res['file_name']}")
                st.write(f"**尺寸 (長*寬*高)**：{res['dimensions_str']}")
                st.write(f"**拆解數值**：長 {res.get('length')} / 寬 {res.get('width')} / 高 {res.get('height')}")
                st.write(f"**單位**：{res['unit']}")
        else:
            st.error(f"❌ {res['file_name']} 解析失敗：{res.get('error_message')}")

    st.markdown("---")
    col_dl, col_rst = st.columns([2, 1])
    
    today_date_str = datetime.now().strftime("%Y%m%d")
    download_filename = f"CAD_Quotation_Report_{today_date_str}{st.session_state.export_ext}"
    
    with col_dl:
        st.download_button(
            label=f"📊 下載完整 CAD 報價單 Excel 檔 ({st.session_state.export_ext})",
            data=st.session_state.excel_bytes,
            file_name=download_filename,
            mime=st.session_state.mime_type,
            use_container_width=True
        )
        
    with col_rst:
        if st.button("🔄 重置 / 準備下一批報價", on_click=reset_session, use_container_width=True):
            st.rerun()
