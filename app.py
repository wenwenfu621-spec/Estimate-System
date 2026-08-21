"""
app.py - Streamlit 網頁介面程式
Version: v2.5.7_20260821
Description: 提供 CAD 報價辨識工具。
             版本升級 v2.5.7：整合 Auto Crop 畫面優化與 Word 4.5 英吋視圖展現。
             自動同步計算 OBB/AABB 並取小值採納，
             展示「加工素材計算採用下列兩種方式之最小值」圖文說明區塊，
             視窗預覽標註尺寸來源模式 (OBB 或 AABB)，
             支援獨立生成與下載 Word 圖文報價單 (.docx) 含等角視圖縮圖，
             含完整 image_error 錯誤診斷碼展現。
             Markdown 顯示轉義星號 (\\.replace("*", "\\*")) 解決顯示問題，
             原生輸入框帶 Tab 切換提示，
             一鍵重置 Widget Key，頁尾含個人頭像徽章 (Design by Max)。
"""

import streamlit as st
import os
import tempfile
import base64
from datetime import datetime
from cad_parser import parse_cad_with_screenshot, generate_excel_report, generate_word_report, is_valid_image


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
    
    /* 示意圖卡片容器樣式 */
    .diagram-card-box {{
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 12px;
        text-align: center;
        margin-bottom: 12px;
    }}
    </style>
    
    <div class="version-badge-left">Version: v2.5.7_20260821</div>
    
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
    """點擊重置按鈕時觸發：清空暫存檔、Session 狀態，並變更 uploader_key 強制清空上傳元件與輸入框"""
    cleanup_temp_files()
    st.session_state.parsed_results = None
    st.session_state.excel_bytes = None
    st.session_state.word_bytes = None
    st.session_state.export_ext = ".xlsx"
    st.session_state.mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    
    st.session_state.uploader_key_num += 1


st.set_page_config(page_title="CAD 報價辨識工具 (v2.5.7)", page_icon="⚙️", layout="centered")

# 載入懸浮元件
inject_custom_elements()

# 標題採用 HTML + white-space: nowrap 鎖定排版，徹底防止齒輪與文字折行堆疊
st.markdown("<h1 style='text-align: center; white-space: nowrap;'>⚙️ CAD 自動報價與尺寸辨識工具 ⚙️</h1>", unsafe_allow_html=True)
st.write("上傳 `.step` 或 `.igs` 3D 模型檔，點選下方按鈕自動辨識尺寸並套寫至 Excel 與 Word 報價單。")

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

# 新增 B4~B7 表頭資訊輸入區塊
with st.expander("📝 客戶與報價表頭資訊填寫 (選填，可直接留空)", expanded=True):
    st.caption("💡 提示：輸入完畢後，按鍵盤 **`Tab`** 鍵可快速切換至下一個輸入欄位。")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        customer_input = st.text_input("客戶名稱", key=f"cust_{st.session_state.uploader_key_num}")
        phone_input = st.text_input("聯絡電話", key=f"phone_{st.session_state.uploader_key_num}")
    with col_c2:
        contact_input = st.text_input("聯絡人", key=f"contact_{st.session_state.uploader_key_num}")
        fax_input = st.text_input("傳真", key=f"fax_{st.session_state.uploader_key_num}")

# 修改說明標題：加工素材計算採用下列兩種方式之最小值
with st.expander("📐 加工素材計算採用下列兩種方式之最小值", expanded=True):
    svg_obb = """
    <svg width="200" height="110" viewBox="0 0 200 110" xmlns="http://www.w3.org/2000/svg">
        <rect width="200" height="110" rx="6" fill="#f8fafc"/>
        <g transform="translate(100, 52) rotate(-20)">
            <rect x="-60" y="-12" width="120" height="24" rx="3" fill="#3b82f6" fill-opacity="0.15" stroke="#2563eb" stroke-width="2" stroke-dasharray="4 2"/>
            <rect x="-55" y="-8" width="110" height="16" rx="4" fill="#64748b" stroke="#334155" stroke-width="1.5"/>
            <!-- 僅保留雙向尺寸箭頭 -->
            <line x1="68" y1="-12" x2="68" y2="12" stroke="#2563eb" stroke-width="2"/>
            <polyline points="65,-8 68,-12 71,-8" fill="none" stroke="#2563eb" stroke-width="2"/>
            <polyline points="65,8 68,12 71,8" fill="none" stroke="#2563eb" stroke-width="2"/>
        </g>
        <text x="100" y="98" font-family="sans-serif" font-size="11" font-weight="bold" fill="#1e293b" text-anchor="middle">最小素材包容盒 (OBB)</text>
    </svg>
    """

    svg_aabb = """
    <svg width="200" height="110" viewBox="0 0 200 110" xmlns="http://www.w3.org/2000/svg">
        <rect width="200" height="110" rx="6" fill="#f8fafc"/>
        <rect x="35" y="18" width="130" height="64" rx="3" fill="#ef4444" fill-opacity="0.12" stroke="#dc2626" stroke-width="2" stroke-dasharray="4 2"/>
        <g transform="translate(100, 50) rotate(-20)">
            <rect x="-55" y="-8" width="110" height="16" rx="4" fill="#64748b" stroke="#334155" stroke-width="1.5"/>
        </g>
        <!-- 僅保留雙向尺寸箭頭 -->
        <line x1="173" y1="18" x2="173" y2="82" stroke="#dc2626" stroke-width="2"/>
        <polyline points="170,22 173,18 176,22" fill="none" stroke="#dc2626" stroke-width="2"/>
        <polyline points="170,78 173,82 176,78" fill="none" stroke="#dc2626" stroke-width="2"/>
        <text x="100" y="98" font-family="sans-serif" font-size="11" font-weight="bold" fill="#1e293b" text-anchor="middle">標準投影外框 (AABB)</text>
    </svg>
    """

    col_img1, col_img2 = st.columns(2)
    with col_img1:
        st.markdown(f'<div class="diagram-card-box">{svg_obb}</div>', unsafe_allow_html=True)
        st.markdown("""
        **【OBB 最小素材包容盒】**
        * 貼合零件幾何方向旋轉量測
        * 素材精確省料 (如傾斜零件估得 4mm)
        * **適用**：CNC 實體備料估價
        """, unsafe_allow_html=True)
    with col_img2:
        st.markdown(f'<div class="diagram-card-box">{svg_aabb}</div>', unsafe_allow_html=True)
        st.markdown("""
        **【AABB 標準投影外框】**
        * 沿世界座標軸 (X/Y/Z) 垂直外包
        * 包含傾斜投影落差 (如傾斜零件估得 9mm)
        * **適用**：外箱體積、正交零件估價
        """, unsafe_allow_html=True)

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

        header_info = {
            "customer": customer_input,
            "contact": contact_input,
            "phone": phone_input,
            "fax": fax_input
        }

        # 1. 產生 Excel 報表
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

        # 2. 產生 Word 圖文報表 (含等角視圖縮圖)
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
    st.subheader("📋 辨識結果預覽")
    for res in st.session_state.parsed_results:
        if res.get("status") == "success":
            # 針對 Markdown 顯示進行星號反斜線轉義 (\*)，並註明採納模式 (OBB/AABB)
            display_dims_escaped = res['dimensions_str'].replace("*", "\\*")
            used_mode = res.get("used_mode", "OBB")
            
            with st.expander(f"📄 {res['file_name']} - 【尺寸：{display_dims_escaped} mm ({used_mode})】"):
                col_p1, col_p2 = st.columns([3, 2])
                with col_p1:
                    st.write(f"**品名 (檔名)**：{res['file_name']}")
                    st.write(f"**尺寸 (長\\*寬\\*高，無條件進位)**：{display_dims_escaped} mm ({used_mode})")
                    st.write(f"**採納演算法**：{used_mode} 模式 (體積較小者)")
                    st.write(f"**拆解數值**：長 {res.get('length')} / 寬 {res.get('width')} / 高 {res.get('height')}")
                    st.write(f"**單位**：{res['unit']}")
                with col_p2:
                    if is_valid_image(res.get("image_path")):
                        st.image(res["image_path"], caption="CAD 等角視圖預覽", use_container_width=True)
                    else:
                        st.warning("⚠️ 等角視圖產生失敗")
                        if res.get("image_error"):
                            st.code(f"Diagnose Error: {res['image_error']}", language="text")

        else:
            st.error(f"❌ {res['file_name']} 解析失敗：{res.get('error_message')}")

    st.markdown("---")
    
    # 雙下載按鈕區域與重置按鈕
    col_dl1, col_dl2, col_rst = st.columns([2, 2, 1])
    
    today_date_str = datetime.now().strftime("%Y%m%d")
    download_excel_name = f"CAD_Quotation_Report_{today_date_str}{st.session_state.export_ext}"
    download_word_name = f"CAD_Quotation_Report_{today_date_str}.docx"
    
    with col_dl1:
        st.download_button(
            label=f"📊 下載 Excel 報價單 ({st.session_state.export_ext})",
            data=st.session_state.excel_bytes,
            file_name=download_excel_name,
            mime=st.session_state.mime_type,
            use_container_width=True
        )

    with col_dl2:
        if st.session_state.word_bytes:
            st.download_button(
                label="📝 下載 CAD 圖文報價單 (.docx)",
                data=st.session_state.word_bytes,
                file_name=download_word_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            st.button("📝 Word 報表不可用", disabled=True, use_container_width=True)
        
    with col_rst:
        if st.button("🔄 重置 / 準備下一批", on_click=reset_session, use_container_width=True):
            st.rerun()
