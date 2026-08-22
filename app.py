import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime
from io import BytesIO, StringIO
from openpyxl import load_workbook

# ==========================================
# CẤU HÌNH GIAO DIỆN (CRYSTAL BAY AIRLINES)
# ==========================================
st.set_page_config(page_title="Crystal Bay Airlines - APIS", page_icon="✈️", layout="wide")

st.markdown("""
<style>
    .stButton > button { 
        border-radius: 10px; font-weight: 600; border: 1px solid #1a2a6c; 
        height: 50px; background-color: white; transition: all 0.3s ease;
    }
    .stButton > button:hover { background-color: #f0f4f8; border: 2px solid #1a2a6c; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# HEADER: HIỂN THỊ 6 NƯỚC + UTC
# ==========================================
header_html = """
<div style="background: linear-gradient(135deg, #1a2a6c, #001f3f); padding: 20px; border-radius: 15px; color: white; font-family: sans-serif; box-shadow: 0 4px 15px rgba(0,0,0,0.3); border-bottom: 4px solid #d4af37;">
    <div style="text-align: center; margin-bottom: 15px;">
        <div style="font-size: 24px; font-weight: 900; letter-spacing: 2px; color: #d4af37;">CRYSTAL BAY AIRLINES</div>
        <div style="font-size: 14px; letter-spacing: 4px; opacity: 0.8;">APIS OPERATIONS CENTER</div>
    </div>
    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; font-size: 12px; text-align: center;">
        <div style="background: rgba(255,255,255,0.08); padding: 5px; border-radius: 8px;">🌐 <b>UTC:</b><br><span id="time-utc" style="color: #d4af37; font-weight: bold;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 5px; border-radius: 8px;">🇻🇳 <b>VN:</b><br><span id="time-vn" style="color: #d4af37; font-weight: bold;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 5px; border-radius: 8px;">🇰🇿 <b>KZ:</b><br><span id="time-kz" style="color: #d4af37; font-weight: bold;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 5px; border-radius: 8px;">🇰🇬 <b>KG:</b><br><span id="time-kg" style="color: #d4af37; font-weight: bold;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 5px; border-radius: 8px;">🇹🇯 <b>TJ:</b><br><span id="time-tj" style="color: #d4af37; font-weight: bold;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 5px; border-radius: 8px;">🇷🇺 <b>RU:</b><br><span id="time-ru" style="color: #d4af37; font-weight: bold;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 5px; border-radius: 8px;">🇵🇱 <b>PL:</b><br><span id="time-pl" style="color: #d4af37; font-weight: bold;"></span></div>
    </div>
</div>
<script>
    function updateTime() {
        const now = new Date();
        const opts = {hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false};
        const zones = ['UTC', 'Asia/Ho_Chi_Minh', 'Asia/Almaty', 'Asia/Bishkek', 'Asia/Dushanbe', 'Europe/Moscow', 'Europe/Warsaw'];
        const ids = ['utc', 'vn', 'kz', 'kg', 'tj', 'ru', 'pl'];
        ids.forEach((id, i) => {
            document.getElementById('time-'+id).innerText = now.toLocaleTimeString('en-GB', {timeZone: zones[i], ...opts});
        });
    }
    setInterval(updateTime, 1000);
    updateTime();
</script>
"""
components.html(header_html, height=250)

# ==========================================
# CẤU HÌNH QUỐC GIA & XỬ LÝ DỮ LIỆU
# ==========================================
COUNTRY_CONFIG = {
    "Việt Nam": "vn", "Kazakhstan": "kz", "Kyrgyzstan": "kg", 
    "Tajikistan": "tj", "Russia": "ru", "Poland": "pl"
}

if 'sel' not in st.session_state: st.session_state.sel = "Việt Nam"

st.markdown("<p style='font-size: 16px; font-weight: bold; color: #1a2a6c;'>🌍 Chọn Quốc gia đến:</p>", unsafe_allow_html=True)
cols = st.columns(3)
for i, name in enumerate(COUNTRY_CONFIG.keys()):
    with cols[i % 3]:
        if st.button(f"{name}", key=f"btn_{name}", use_container_width=True):
            st.session_state.sel = name
            st.rerun()

st.markdown("---")
uploaded_gd = st.file_uploader(f"Tải file GD cho chuyến bay đến {st.session_state.sel}", type=["xls", "xlsx"])

if uploaded_gd:
    st.info(f"Đang chờ xử lý dữ liệu cho: **{st.session_state.sel}**")
