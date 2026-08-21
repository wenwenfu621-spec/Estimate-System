"""
pages/2_2D_Mylar.py - 2D Mylar / DXF 材料辨識專屬頁面
Version: v2.8.3_20260821
"""

import streamlit as st
import os
import sys
import tempfile
from datetime import datetime

# --- 修復模組路徑 ---
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from dxf_parser import parse_dxf_2d
from cad_parser import generate_excel_report, generate_word_report, is_valid_image

def cleanup_2d_temp():
    if "temp_files_2d" in st.session_state and st.session_state.temp_files_2d:
        for fpath in st.session_state.temp_files_2d:
            if fpath and os.path.exists(fpath):
                try: os.remove(fpath)
                except: pass
        st.session_state.temp_files_2d = []

def reset_2d_session():
    cleanup_2d_temp()
    st.session_state.parsed_2d_results = None
    st.session_state.excel_2d_bytes = None
    st.session_state.uploader_2d_key += 1

st.set_page_config(page_title="2D Mylar 材料辨識", page_icon="📄", layout="centered")

if st.button("← 返回功能首頁"):
    st.switch_page("app.py")

st.title("📄 2D Mylar / 模切材料辨識")

# --- 設定參數 ---
material_type = st.selectbox("材料類型", ["Mylar", "PET", "PC Film", "Other"])
selected_thickness = st.number_input("厚度 (mm)", value=0.25, step=0.01)

uploaded_2d_files = st.file_uploader("上傳 DXF", type=["dxf"], accept_multiple_files=True)

if uploaded_2d_files and st.button("🎯 開始辨識", type="primary", width="stretch"):
    cleanup_2d_temp()
    parsed_results = []
    for up_file in uploaded_2d_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
            tmp.write(up_file.getvalue())
            res = parse_dxf_2d(tmp.name, user_thickness=selected_thickness)
            res["file_name"] = up_file.name
            res["material_type"] = material_type
            parsed_results.append(res)
            st.session_state.temp_files_2d.append(tmp.name)

    has_xlsm = os.path.exists("template.xlsm") or os.path.exists("Template.xlsm")
    export_ext = ".xlsm" if has_xlsm else ".xlsx"
    excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=export_ext)
    generate_excel_report(parsed_results, excel_tmp.name)
    with open(excel_tmp.name, "rb") as f: st.session_state.excel_2d_bytes = f.read()
    st.session_state.parsed_2d_results = parsed_results
    st.session_state.temp_files_2d.append(excel_tmp.name)

if st.session_state.parsed_2d_results:
    st.write("辨識完成！")
    today_str = datetime.now().strftime("%Y%m%d")
    has_xlsm = os.path.exists("template.xlsm") or os.path.exists("Template.xlsm")
    d_ext, d_mime = (".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12") if has_xlsm else (".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    st.download_button("📊 下載 Excel", st.session_state.excel_2d_bytes, f"2D_Report_{today_str}{d_ext}", mime=d_mime, width="stretch")
    st.button("🔄 重置頁面", on_click=reset_2d_session, width="stretch")
