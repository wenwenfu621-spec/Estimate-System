"""
app.py - Streamlit 系統首頁 (Landing Page)
Version: v2.8.3_20260821
Description: 提供 3D CAD 與 2D Mylar 雙軌功能獨立入口。
             修正：switch_page 路徑已對齊 GitHub 真實檔名。
"""

import streamlit as st
import os
import base64


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
    .landing-card {{
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 24px; text-align: center; margin-bottom: 16px; box-shadow: 0px 2px 4px rgba(0,0,0,0.05);
        height: 220px; display: flex; flex-direction: column; justify-content: space-between;
    }}
    </style>
    <div class="version-badge-left">Version: v2.8.3_20260821</div>
    <div class="custom-footer-max">{avatar_html}<span class="custom-footer-text">Design by Max</span></div>
    """
    st.markdown(custom_css, unsafe_allow_html=True)


st.set_page_config(page_title="CAD 與 2D 模切自動報價系統", page_icon="⚙️", layout="centered")
inject_custom_elements()

st.markdown("<h1 style='text-align: center; white-space: nowrap;'>⚙️ CAD 自動報價與尺寸辨識系統 ⚙️</h1>", unsafe_allow_html=True)
st.write("")
st.markdown("<h3 style='text-align: center; color: #475569;'>請選擇您要執行的估價辨識功能：</h3>", unsafe_allow_html=True)
st.write("")

col1, col2 = st.columns(2, gap="medium")

with col1:
    st.markdown("""
    <div class="landing-card">
        <div>
            <h3>📦 3D 機構件尺寸辨識</h3>
            <p style="color: #64748b; font-size: 0.85rem;">適用 STEP / STP / IGES / IGS<br>具備 OBB/AABB 最小包容盒、工程三視圖與 Excel/Word 報價匯出。</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("進入 3D CAD 辨識", width="stretch", type="primary"):
        # 確保此處與 GitHub 上的 1_3D_CAD.py 大小寫完全一致
        st.switch_page("pages/1_3D_CAD.py")

with col2:
    st.markdown("""
    <div class="landing-card">
        <div>
            <h3>📄 2D Mylar 材料辨識</h3>
            <p style="color: #64748b; font-size: 0.85rem;">適用 DXF 2D 模切檔案<br>具備外形輪廓尺寸、材料厚度設定、面積計算與 2D 預覽。</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("進入 2D Mylar 辨識", width="stretch", type="primary"):
        # ⚠️ 這裡修改為與您截圖中完全一致的檔名：2_2D_MYLAR.py (全大寫)
        st.switch_page("pages/2_2D_MYLAR.py")
