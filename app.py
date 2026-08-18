"""
app.py - Streamlit 網頁介面程式
Version: v1.4.1_20260818
Description: 提供多檔 CAD 上傳、按鈕觸發辨識、預覽尺寸與截圖，支援匯出 Excel，
             並修正個人識別頭像檔名大小寫匹配與 HTML 渲染問題 (Design by Max)。
"""

import streamlit as st
import os
import tempfile
import base64
from cad_parser import parse_cad_with_screenshot, generate_excel_report


def inject_custom_footer():
    """於頁面下方注入個人識別頭像與 Design by Max 徽章 (支援大小寫檔名相容)"""
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

    footer_css = f"""
    <style>
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
    <div class="custom-footer-max">
        {avatar_html}
        <span class="custom-footer-text">Design by Max</span>
    </div>
    """
    st.markdown(footer_css, unsafe_allow_html=True)


st.set_page_config(page_title="CAD 報價辨識工具 (v1.4.1)", page_icon="⚙️", layout="centered")

st.title("⚙️ CAD 自動報價與尺寸辨識工具")
st.caption("版本別：`v1.4.1_20260818` (頭像修正版 + 個人識別徽章)")
st.write("上傳 `.step` 或 `.igs` 3D 模型檔，點選下方按鈕自動辨識尺寸與擷取 3D 視角圖像。")

uploaded_files = st.file_uploader(
    "上傳 CAD 圖檔 (可多選)", 
    type=["step", "stp", "igs", "iges"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"已選擇 **{len(uploaded_files)}** 個檔案。")

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

            result = parse_cad_with_screenshot(tmp_path)
            result["file_name"] = uploaded_file.name
            parsed_results.append(result)

            if os.path.exists(tmp_path):
                os.remove(tmp_path)

            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.success("🎉 所有檔案辨識完成！")

        st.subheader("📋 辨識結果預覽")
        for res in parsed_results:
            if res.get("status") == "success":
                with st.expander(f"📄 {res['file_name']} - 【{res['dimensions_str']} mm】"):
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        img_p = res.get("image_path")
                        if img_p and os.path.exists(img_p):
                            st.image(img_p, caption="3D 視角縮圖", width=150)
                        else:
                            st.info("無 3D 預覽圖")
                    with col2:
                        st.write(f"**檔名**：{res['file_name']}")
                        st.write(f"**長寬高**：{res['dimensions_str']}")
                        st.write(f"**單位**：{res['unit']}")
            else:
                st.error(f"❌ {res['file_name']} 解析失敗：{res.get('error_message')}")

        excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        excel_path = excel_tmp.name
        excel_tmp.close()

        generate_excel_report(parsed_results, excel_path)

        with open(excel_path, "rb") as f:
            excel_bytes = f.read()

        st.markdown("---")
        st.download_button(
            label="📊 下載 CAD 辨識結果 Excel 報表 (.xlsx)",
            data=excel_bytes,
            file_name="CAD_Quotation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if os.path.exists(excel_path):
            os.remove(excel_path)
        for res in parsed_results:
            img_p = res.get("image_path")
            if img_p and os.path.exists(img_p):
                os.remove(img_p)

# 載入頁尾個人識別頭像徽章
inject_custom_footer()
