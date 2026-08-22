import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO, StringIO

def parse_date(date_str):
    if pd.isna(date_str) or not str(date_str).strip() or str(date_str).strip().lower() == 'nan':
        return ""
    date_str = str(date_str).strip()
    try:
        d = datetime.strptime(date_str, "%d/%m/%y")
        if d.year > 2050:
            d = d.replace(year=d.year - 100)
        return d.strftime("%d/%m/%Y")
    except:
        return date_str

def process_roster_data(gd_file, template_file_path):
    # 1. Đọc file GD (Giữ nguyên logic cũ đã tối ưu)
    content = gd_file.getvalue()
    df_gd = None
    try:
        df_gd = pd.read_excel(BytesIO(content))
    except:
        for enc in ['utf-8', 'latin1', 'cp1258']:
            try:
                dfs = pd.read_html(StringIO(content.decode(enc, errors='ignore')))
                if dfs: df_gd = dfs[0]; break
            except: continue
    
    if df_gd is None: raise ValueError("Không thể đọc file GD.")

    # 2. Tìm bảng dữ liệu
    header_idx = None
    for idx, row in df_gd.iterrows():
        row_str = row.fillna("").astype(str).str.lower()
        if any('passport' in s for s in row_str) and any('name' in s for s in row_str):
            header_idx = idx
            break
    
    # 3. Trích xuất dữ liệu
    header_row = df_gd.iloc[header_idx].fillna("").astype(str).str.lower()
    col_name = next((i for i, v in enumerate(header_row) if 'name' in v), None)
    col_passport = next((i for i, v in enumerate(header_row) if 'passport' in v and 'expiry' not in v), None)
    col_dob = next((i for i, v in enumerate(header_row) if 'birth' in v), None)
    col_gender = next((i for i, v in enumerate(header_row) if v == 'g'), None)
    col_nat = next((i for i, v in enumerate(header_row) if 'ntly' in v), None)
    col_expiry = next((i for i, v in enumerate(header_row) if 'expiry' in v), None)

    crew_data = []
    for idx in range(header_idx + 1, len(df_gd)):
        row = df_gd.iloc[idx]
        if 'declaration of health' in row.fillna("").astype(str).str.cat(sep=" ").lower(): break
        
        name_val = str(row.iloc[col_name]).strip() if col_name is not None and pd.notna(row.iloc[col_name]) else 'nan'
        passport_val = str(row.iloc[col_passport]).strip() if col_passport is not None and pd.notna(row.iloc[col_passport]) else 'nan'
        
        if name_val.lower() != 'nan' and passport_val.lower() != 'nan':
            name_parts = name_val.split()
            if len(name_parts) > 1 and len(name_parts[-1]) <= 3 and name_parts[-1].isupper(): name_parts = name_parts[:-1]
            family = name_parts[0] if name_parts else ""; given = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
            nat = str(row.iloc[col_nat]).strip() if col_nat is not None and pd.notna(row.iloc[col_nat]) else ""
            crew_data.append([family, np.nan, given, str(row.iloc[col_gender]).strip() if col_gender is not None else "", 
                              'VN' if nat.upper() == 'VNM' else nat, parse_date(row.iloc[col_dob]) if col_dob is not None else "", 
                              'P', passport_val, 'VN' if nat.upper() == 'VNM' else nat, parse_date(row.iloc[col_expiry]) if col_expiry is not None else ""])
            
    # 4. Ghi đè vào file mẫu mà vẫn giữ nguyên định dạng
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Load file mẫu bằng openpyxl để giữ nguyên định dạng (font, màu, border)
        from openpyxl import load_workbook
        book = load_workbook(template_file_path)
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}
        
        # Giả sử bảng bắt đầu từ dòng 13 (index 12)
        start_row = 13 
        sheet = book.active
        for r_idx, row_data in enumerate(crew_data, start_row):
            for c_idx, value in enumerate(row_data, 1):
                sheet.cell(row=r_idx, column=c_idx, value=value)
        book.save(output)
        
    return output.getvalue()

st.set_page_config(page_title="Roster Pro", page_icon="✈️")
st.title("✈️ Roster Tool: Chuyển đổi GD sang VNAPIS")
uploaded_gd = st.file_uploader("Tải lên file GD (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_gd:
    try:
        excel_data = process_roster_data(uploaded_gd, "Template_VNAPIS.xlsx")
        st.success("✅ Đã xử lý xong!")
        st.download_button("⬇️ Tải file kết quả (Giữ nguyên định dạng)", excel_data, "VNAPIS_Final.xlsx")
    except Exception as e: st.error(f"Lỗi: {e}")
