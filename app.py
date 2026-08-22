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
        border-radius: 10px; 
        font-weight: 600; 
        border: 1px solid #1a2a6c; 
        height: 60px; 
        background-color: white; 
        transition: all 0.3s ease;
    }
    .stButton > button:hover { 
        background-color: #f0f4f8; 
        border: 2px solid #1a2a6c; 
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(26,42,108,0.2);
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# HEADER: HIỂN THỊ ĐỒNG HỒ 6 MÚI GIỜ (CRYSTAL BAY)
# ==========================================
header_html = """
<div style="background: linear-gradient(135deg, #1a2a6c, #001f3f); padding: 20px; border-radius: 15px; color: white; font-family: sans-serif; box-shadow: 0 4px 15px rgba(0,0,0,0.3); border-bottom: 4px solid #d4af37;">
    <div style="text-align: center; margin-bottom: 15px;">
        <div style="font-size: 24px; font-weight: 900; letter-spacing: 2px; color: #d4af37;">CRYSTAL BAY AIRLINES</div>
        <div style="font-size: 14px; letter-spacing: 4px; opacity: 0.8;">APIS CENTER</div>
    </div>
    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; font-size: 13px; text-align: center;">
        <div style="background: rgba(255,255,255,0.08); padding: 8px; border-radius: 8px;"><b> VietNam (VN):</b><br><span id="time-vn" style="font-size: 16px; font-weight: bold; color: #d4af37;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 8px; border-radius: 8px;"><b>🌐 UTC:</b><br><span id="time-utc" style="font-size: 16px; font-weight: bold; color: #d4af37;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 8px; border-radius: 8px;"><b> Kazakhstan (KZ):</b><br><span id="time-kz" style="font-size: 16px; font-weight: bold; color: #d4af37;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 8px; border-radius: 8px;"><b> Kyrgyzstan (KG):</b><br><span id="time-kg" style="font-size: 16px; font-weight: bold; color: #d4af37;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 8px; border-radius: 8px;"><b> Tajikistan (TJ):</b><br><span id="time-tj" style="font-size: 16px; font-weight: bold; color: #d4af37;"></span></div>
        <div style="background: rgba(255,255,255,0.08); padding: 8px; border-radius: 8px;"><b> Russia (RU):</b><br><span id="time-ru" style="font-size: 16px; font-weight: bold; color: #d4af37;"></span></div>
    </div>
</div>
<script>
    function updateTime() {
        const now = new Date();
        const opts = {hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false};
        const zones = ['Asia/Ho_Chi_Minh', 'UTC', 'Asia/Almaty', 'Asia/Bishkek', 'Asia/Dushanbe', 'Europe/Moscow'];
        const ids = ['vn', 'utc', 'kz', 'kg', 'tj', 'ru'];
        ids.forEach((id, i) => {
            document.getElementById('time-'+id).innerText = now.toLocaleTimeString('en-GB', {timeZone: zones[i], ...opts});
        });
    }
    setInterval(updateTime, 1000);
    updateTime();
</script>
"""
components.html(header_html, height=220)

# ==========================================
# CÁC HÀM XỬ LÝ DỮ LIỆU THÔNG MINH
# ==========================================
def parse_date(date_str):
    if pd.isna(date_str) or not str(date_str).strip() or str(date_str).strip().lower() == 'nan': return ""
    date_str = str(date_str).strip()
    try:
        d = datetime.strptime(date_str, "%d/%m/%y")
        if d.year > 2050: d = d.replace(year=d.year - 100)
        return d.strftime("%d/%m/%Y")
    except: return date_str

def process_roster_data_vn(gd_file, template_file_path):
    content = gd_file.getvalue()
    df_gd = None
    
    # Đọc file đa định dạng
    for engine in ['openpyxl', 'xlrd']:
        try: 
            df_gd = pd.read_excel(BytesIO(content), engine=engine)
            if df_gd is not None and len(df_gd) > 0: break
        except: continue

    if df_gd is None or len(df_gd) == 0:
        for enc in ['utf-8', 'latin1', 'cp1258']:
            try:
                dfs = pd.read_html(StringIO(content.decode(enc, errors='ignore')))
                if len(dfs) > 0: 
                    df_gd = dfs[0]
                    break
            except: continue

    if df_gd is None or len(df_gd) == 0: 
        raise ValueError("Không thể đọc được dữ liệu trong file GD. Định dạng file không được hỗ trợ.")

    # 1. TỰ ĐỘNG DÒ TÌM NƠI KHỞI HÀNH & NƠI ĐẾN CHÍNH XÁC
    dep_place, arr_place = "", ""
    for idx in range(min(15, len(df_gd))):
        row_full_text = " ".join(df_gd.iloc[idx].fillna("").astype(str))
        row_lower = row_full_text.lower()
        
        if "departure from" in row_lower:
            try:
                part = row_full_text.split("Departure from")[1].strip()
                dep_place = part.split()[0].upper()
            except:
                pass
                
        if "arrival at" in row_lower:
            try:
                part = row_full_text.split("Arrival at")[1].strip()
                arr_place = part.split()[0].upper()
            except:
                pass

    # 2. TÌM HEADER BẢNG TỔ BAY
    header_idx = None
    for idx, row in df_gd.iterrows():
        row_str = row.fillna("").astype(str).str.lower()
        if any('passport' in s for s in row_str) and any('name' in s for s in row_str):
            header_idx = idx
            break
            
    if header_idx is None: 
        raise ValueError("Không tìm thấy bảng danh sách tổ bay (thiếu cột 'passport' hoặc 'name') trong file GD.")
        
    header_row = df_gd.iloc[header_idx].fillna("").astype(str).str.lower()
    col_name = next((i for i, v in enumerate(header_row) if 'name' in v), None)
    col_passport = next((i for i, v in enumerate(header_row) if 'passport' in v and 'expiry' not in v), None)
    col_dob = next((i for i, v in enumerate(header_row) if 'birth' in v), None)
    col_gender = next((i for i, v in enumerate(header_row) if v == 'g'), None)
    col_nat = next((i for i, v in enumerate(header_row) if 'ntly' in v), None)
    col_expiry = next((i for i, v in enumerate(header_row) if 'expiry' in v), None)

    crew_data = []
    seen_passports = set()
    
    for idx in range(header_idx + 1, len(df_gd)):
        row = df_gd.iloc[idx]
        row_text = row.fillna("").astype(str).str.cat(sep=" ").lower()
        if 'declaration of health' in row_text or 'total' in row_text: 
            break
            
        name_val = str(row.iloc[col_name]).strip() if col_name is not None and pd.notna(row.iloc[col_name]) else 'nan'
        passport_val = str(row.iloc[col_passport]).strip() if col_passport is not None and pd.notna(row.iloc[col_passport]) else 'nan'
        
        if name_val.lower() != 'nan' and passport_val.lower() != 'nan' and name_val != '':
            if passport_val in seen_passports: 
                continue
            seen_passports.add(passport_val)
            
            name_parts = name_val.split()
            if len(name_parts) > 1 and len(name_parts[-1]) <= 3 and name_parts[-1].isupper(): 
                name_parts = name_parts[:-1]
                
            family_name = name_parts[0] if name_parts else ""
            given_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
            nat = str(row.iloc[col_nat]).strip() if col_nat is not None and pd.notna(row.iloc[col_nat]) else ""
            if nat.upper() == 'VNM': 
                nat = 'VN'
            gender = str(row.iloc[col_gender]).strip() if col_gender is not None and pd.notna(row.iloc[col_gender]) else ""
            
            crew_data.append([
                family_name, None, given_name, gender, nat, 
                parse_date(row.iloc[col_dob]) if col_dob is not None else "", 
                'P', passport_val, nat, 
                parse_date(row.iloc[col_expiry]) if col_expiry is not None else ""
            ])

    if len(crew_data) == 0:
        raise ValueError("Đọc được file nhưng không tìm thấy dữ liệu tổ bay hợp lệ bên trong.")

    # 3. GHI DỮ LIỆU VÀO TEMPLATE
    output = BytesIO()
    book = load_workbook(template_file_path)
    sheet = book.active
    
    if dep_place: sheet['B3'] = dep_place
    if arr_place: sheet['B6'] = arr_place
    
    for row in sheet.iter_rows(min_row=14, max_row=sheet.max_row, min_col=1, max_col=10):
        for cell in row: cell.value = None

    for r_idx, row_data in enumerate(crew_data, 14):
        for c_idx, value in enumerate(row_data, 1):
            sheet.cell(row=r_idx, column=c_idx, value=value)
            
    book.save(output)
    df_preview = pd.DataFrame(crew_data, columns=['Họ', 'Tên đệm', 'Tên', 'Giới tính', 'Quốc tịch', 'Ngày sinh', 'Loại giấy tờ', 'Số giấy tờ', 'Nơi cấp', 'Ngày hết hạn'])
    return output.getvalue(), df_preview

# ==========================================
# GIAO DIỆN CHỌN 06 QUỐC GIA (CRYSTAL BAY)
# ==========================================
COUNTRY_CONFIG = {
    "Việt Nam": {"code": "vn", "ready": True, "template": "Template_VNAPIS.xlsx"},
    "Kazakhstan": {"code": "kz", "ready": False, "template": ""},
    "Kyrgyzstan": {"code": "kg", "ready": False, "template": ""},
    "Tajikistan": {"code": "tj", "ready": False, "template": ""},
    "Russia": {"code": "ru", "ready": False, "template": ""},
    "Poland": {"code": "pl", "ready": False, "template": ""}
}

if 'sel' not in st.session_state: 
    st.session_state.sel = "Việt Nam"

st.markdown("<p style='font-size: 16px; font-weight: bold; color: #1a2a6c; margin-top: 20px;'>🌍 Vui lòng chọn Quốc gia đến:</p>", unsafe_allow_html=True)

keys = list(COUNTRY_CONFIG.keys())

def render_country_grid(start_idx, end_idx):
    cols = st.columns(3) # 3 cột mỗi hàng cho cân đối 6 nước
    for i, idx in enumerate(range(start_idx, end_idx)):
        c_name = keys[idx]
        cfg = COUNTRY_CONFIG[c_name]
        is_selected = (st.session_state.sel == c_name)
        
        with cols[i]:
            flag_url = f"https://flagcdn.com/w40/{cfg['code']}.png"
            border_color = "#1a2a6c" if is_selected else "#d1d5db"
            bg_color = "#f0f4f8" if is_selected else "#ffffff"
            shadow = "0 4px 10px rgba(26,42,108,0.2)" if is_selected else "0 2px 4px rgba(0,0,0,0.05)"
            
            st.markdown(f"""
            <div style="display: flex; align-items: center; padding: 8px 12px; background-color: {bg_color}; border: 2px solid {border_color}; border-radius: 8px; box-shadow: {shadow}; margin-bottom: 6px;">
                <img src="{flag_url}" width="32" style="border-radius: 3px; margin-right: 10px; object-fit: cover;">
                <div style="font-weight: 700; color: #1a2a6c; font-size: 13px;">{c_name}</div>
            </div>
            """, unsafe_allow_html=True)
            
            btn_text = f"Chọn {c_name}" if not is_selected else "Đang chọn ✓"
            if st.button(btn_text, key=f"btn_{c_name}", use_container_width=True):
                st.session_state.sel = c_name
                st.rerun()

render_country_grid(0, 3)
render_country_grid(3, 6)

st.markdown("---")

# ==========================================
# XỬ LÝ UPLOAD VÀ XUẤT FILE APIS
# ==========================================
selected_cfg = COUNTRY_CONFIG[st.session_state.sel]

if selected_cfg["ready"]:
    uploaded_gd = st.file_uploader(f"Tải lên file GD (.xls, .xlsx) cho chuyến bay đến {st.session_state.sel}", type=["xls", "xlsx", "txt", "csv"])
    if uploaded_gd is not None:
        st.info(f"Đang xử lý dữ liệu APIS cho {st.session_state.sel}...")
        try:
            if st.session_state.sel == "Việt Nam":
                excel_data, preview_data = process_roster_data_vn(uploaded_gd, selected_cfg["template"])
                
            st.success(f"✅ Đã xử lý thành công form APIS cho {st.session_state.sel}!")
            st.dataframe(preview_data, use_container_width=True) 
            st.download_button(
                label=f"⬇️ Tải file Excel APIS chuẩn ({st.session_state.sel})",
                data=excel_data,
                file_name=f"APIS_Crew_{selected_cfg['code'].upper()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý: {e}")
else:
    st.warning(f"🚧 Chức năng xuất APIS cho **{st.session_state.sel}** đang được xây dựng (chờ template chuẩn).")
