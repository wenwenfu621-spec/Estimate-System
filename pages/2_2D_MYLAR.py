"""
pages/2_2D_MYLAR.py - 2D Mylar / 模切材料快速報價頁面
Version: v2.8.4_20260821
Description: 支援 PDF / DXF 混合上傳，具備 PDF 尺寸提取、人工確認與手動 Fallback 機制。
"""

import streamlit as st
import os
import sys
import tempfile
from datetime import datetime

# --- 修復模組路徑 ---
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dxf_parser import parse_dxf_2d
from pdf_2d_parser import parse_pdf_2d
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


st.set_page_config(page_title="2D Mylar 快速報價", page_icon="📄", layout="centered")

if st.button("← 返回功能首頁"):
    st.switch_page("app.py")

st.title("📄 2D Mylar / 模切材料快速報價系統")
st.write("上傳 `.pdf` 或 `.dxf` 2D 圖檔（可多選），系統自動擷取外形尺寸以進行快速報價。")

if "uploader_2d_key" not in st.session_state:
    st.session_state.uploader_2d_key = 0
if "parsed_2d_results" not in st.session_state:
    st.session_state.parsed_2d_results = None
if "excel_2d_bytes" not in st.session_state:
    st.session_state.excel_2d_bytes = None
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

# 驗證 Thickness > 0
if selected_thickness <= 0:
    st.error("⚠️ 請輸入有效的材料厚度（必須大於 0）")
    st.stop()

st.info("⚠️ **DWG 提示**：DWG 目前不直接解析；正式製作請提供 DXF / DWG 工程資料。")

uploaded_2d_files = st.file_uploader(
    "上傳 2D 模切圖檔（可多選）\n\n支援格式：PDF / DXF", 
    type=["pdf", "dxf"],
    accept_multiple_files=True,
    key=f"uploader_2d_{st.session_state.uploader_2d_key}"
)

if uploaded_2d_files:
    st.write(f"已選擇 **{len(uploaded_2d_files)}** 個 2D 檔案。")

    if st.button("🎯 開始 2D 尺寸快速辨識", type="primary", width="stretch"):
        cleanup_2d_temp()
        parsed_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, up_file in enumerate(uploaded_2d_files):
            fname = up_file.name
            ext = os.path.splitext(fname)[1].lower()
            status_text.text(f"正在處理檔案 ({idx+1}/{len(uploaded_2d_files)}): {fname} ...")

            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                tmp.write(up_file.getvalue())
                tmp_path = tmp.name
                st.session_state.temp_files_2d.append(tmp_path)

            if ext == '.dxf':
                res = parse_dxf_2d(tmp_path, user_thickness=selected_thickness)
                res["file_name"] = fname
                res["source_type"] = "DXF"
                res["dimension_source"] = "DXF Geometry"
                res["confirmed"] = True
                parsed_results.append(res)
            elif ext == '.pdf':
                res = parse_pdf_2d(tmp_path)
                res["file_name"] = fname
                res["source_type"] = "PDF"
                res["thickness"] = selected_thickness
                res["material_type"] = material_type
                if res.get("status") == "success":
                    l_val = max(res.get("length", 0), res.get("width", 0))
                    w_val = min(res.get("length", 0), res.get("width", 0))
                    res["length"] = l_val
                    res["width"] = w_val
                    res["dimensions_str"] = f"{l_val}*{w_val}"
                    res["confirmed"] = False  # 需要人工確認
                parsed_results.append(res)

            progress_bar.progress((idx + 1) / len(uploaded_2d_files))

        status_text.success("🎉 2D 檔案初步辨識完成！請在下方進行確認或手動調整。")
        st.session_state.parsed_2d_results = parsed_results

if st.session_state.parsed_2d_results:
    st.subheader("📋 2D 快速報價尺寸確認與調整")
    
    updated_results = []
    all_confirmed = True

    for i, res in enumerate(st.session_state.parsed_2d_results):
        fname = res.get("file_name", f"Item_{i}")
        st.markdown(f"---")
        st.markdown(f"### 檔案：`{fname}` ({res.get('source_type', 'Unknown')})")

        if res.get("source_type") == "PDF":
            st.info("ℹ️ **PDF Dimension / Quotation Reference Only**\n本尺寸依客戶提供 PDF 圖面標註進行估價；正式製作尺寸以經確認之 DXF / DWG 工程資料為準。")

        if res.get("status") == "success":
            default_l = float(res.get("length", 100.0))
            default_w = float(res.get("width", 50.0))

            col_u1, col_u2, col_u3 = st.columns(3)
            with col_u1:
                user_l = st.number_input(f"Length (mm) [{i}]", min_value=0.1, value=default_l, step=1.0, key=f"l_{i}")
            with col_u2:
                user_w = st.number_input(f"Width (mm) [{i}]", min_value=0.1, value=default_w, step=1.0, key=f"w_{i}")
            with col_u3:
                is_conf = st.checkbox(f"確認採用此尺寸 [{i}]", value=res.get("confirmed", False), key=f"conf_{i}")

            # L >= W 正規化
            final_l = max(user_l, user_w)
            final_w = min(user_l, user_w)
            gross_area = final_l * final_w

            res["length"] = final_l
            res["width"] = final_w
            res["thickness"] = selected_thickness
            res["material_type"] = material_type
            res["dimensions_str"] = f"{final_l}*{final_w}"
            res["gross_area"] = gross_area
            res["net_area"] = gross_area
            res["confirmed"] = is_conf

            if not is_conf:
                all_confirmed = False
        else:
            # 發生錯誤時提供手動 Fallback
            st.warning(f"⚠️ 自動尺寸辨識失敗 ({res.get('error_message', '未知錯誤')})，請手動輸入尺寸。")
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                manual_l = st.number_input(f"手動 Length (mm) [{i}]", min_value=0.1, value=100.0, step=1.0, key=f"ml_{i}")
            with col_f2:
                manual_w = st.number_input(f"手動 Width (mm) [{i}]", min_value=0.1, value=50.0, step=1.0, key=f"mw_{i}")
            with col_f3:
                manual_conf = st.checkbox(f"確認並採用手動尺寸 [{i}]", value=True, key=f"mconf_{i}")

            final_l = max(manual_l, manual_w)
            final_w = min(manual_l, manual_w)
            gross_area = final_l * final_w

            res["status"] = "success"
            res["length"] = final_l
            res["width"] = final_w
            res["thickness"] = selected_thickness
            res["material_type"] = material_type
            res["dimensions_str"] = f"{final_l}*{final_w}"
            res["gross_area"] = gross_area
            res["net_area"] = gross_area
            res["confirmed"] = manual_conf

            if not manual_conf:
                all_confirmed = False

        updated_results.append(res)

    st.markdown("---")
    if all_confirmed:
        header_info = {"customer": customer_input, "contact": contact_input, "phone": phone_input, "fax": fax_input}

        has_xlsm = os.path.exists("template.xlsm") or os.path.exists("Template.xlsm")
        export_ext = ".xlsm" if has_xlsm else ".xlsx"

        excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=export_ext)
        excel_path = excel_tmp.name
        excel_tmp.close()
        st.session_state.temp_files_2d.append(excel_path)

        generate_excel_report(updated_results, excel_path, header_info=header_info)
        with open(excel_path, "rb") as f:
            excel_bytes = f.read()

        today_str = datetime.now().strftime("%Y%m%d")
        download_ext = ".xlsm" if has_xlsm else ".xlsx"
        download_mime = "application/vnd.ms-excel.sheet.macroEnabled.12" if has_xlsm else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.download_button(
                "📊 下載 2D 快速報價 Excel", 
                excel_bytes, 
                f"2D_Quotation_Report_{today_str}{download_ext}", 
                mime=download_mime, 
                width="stretch"
            )
        with col_b2:
            st.button("🔄 重置 2D 頁面", on_click=reset_2d_session, width="stretch")
    else:
        st.warning("⚠️ 請確認並勾選所有檔案的「確認採用」後，即可解鎖 Excel 報價單下載。")
