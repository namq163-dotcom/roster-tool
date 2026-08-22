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
    # Đọc file GD linh hoạt (hỗ trợ cả Excel và file HTML định dạng .xls)
    df_gd = None
    try:
        df_gd = pd.read_excel(gd_file)
    except Exception:
        try:
            content = gd_file.getvalue()
            # Thử giải mã bằng utf-8 hoặc latin1 nếu file HTML tiếng Việt
            try:
                html_content = content.decode('utf-8')
            except:
                html_content = content.decode('latin1', errors='ignore')
            
            dfs = pd.read_html(StringIO(html_content))
            df_gd = dfs[0]
        except Exception as e:
            raise ValueError(f"Không thể đọc file GD. Chi tiết: {e}")
    
    header_idx = None
    for idx, row in df_gd.iterrows():
        row_str = row.astype(str).str.lower()
        if any('passport' in s for s in row_str) and any('name' in s for s in row_str):
            header_idx = idx
            break
            
    if header_idx is None:
        raise ValueError("Không tìm thấy bảng danh sách tổ bay trong file GD.")
        
    header_row = df_gd.iloc[header_idx].astype(str).str.lower()
    col_name = col_passport = col_dob = col_gender = col_nat = col_expiry = None
    
    for idx, val in enumerate(header_row):
        if 'name' in val: col_name = idx
        elif 'passport' in val and 'expiry' not in val: col_passport = idx
        elif 'birth' in val: col_dob = idx
        elif val == 'g': col_gender = idx
        elif val == 'ntly': col_nat = idx
        elif 'expiry' in val: col_expiry = idx

    crew_data = []
    for idx in range(header_idx + 1, len(df_gd)):
        row = df_gd.iloc[idx]
        
        if row.astype(str).str.contains('DECLARATION OF HEALTH', case=False, na=False).any():
            break
            
        name_val = str(row.iloc[col_name]).strip() if pd.notna(row.iloc[col_name]) else 'nan'
        passport_val = str(row.iloc[col_passport]).strip() if pd.notna(row.iloc[col_passport]) else 'nan'
        
        if name_val.lower() != 'nan' and passport_val.lower() != 'nan':
            name_parts = name_val.split()
            if len(name_parts) > 1 and len(name_parts[-1]) <= 3 and name_parts[-1].isupper():
                name_parts = name_parts[:-1]
                
            family_name = name_parts[0] if name_parts else ""
            given_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
            
            nat = str(row.iloc[col_nat]).strip() if pd.notna(row.iloc[col_nat]) else ""
            if nat == 'VNM': nat = 'VN'
            
            crew_data.append({
                'Họ': family_name,
                'Tên đệm': np.nan, 
                'Tên': given_name,
                'Giới tính': str(row.iloc[col_gender]).strip(),
                'Quốc tịch': nat,
                'Ngày sinh': parse_date(row.iloc[col_dob]),
                'Loại giấy tờ': 'P', 
                'Số giấy tờ': passport_val,
                'Nơi cấp': nat, 
                'Ngày hết hạn': parse_date(row.iloc[col_expiry])
            })
            
    df_crew = pd.DataFrame(crew_data)
    
    df_template = pd.read_excel(template_file_path, sheet_name=0, header=None)
    header_template = df_template.iloc[:12].copy()
    
    df_formatted = pd.DataFrame()
    df_formatted[0] = df_crew['Họ']
    df_formatted[1] = df_crew['Tên đệm']
    df_formatted[2] = df_crew['Tên']
    df_formatted[3] = df_crew['Giới tính']
    df_formatted[4] = df_crew['Quốc tịch']
    df_formatted[5] = df_crew['Ngày sinh']
    df_formatted[6] = df_crew['Loại giấy tờ']
    df_formatted[7] = df_crew['Số giấy tờ']
    df_formatted[8] = df_crew['Nơi cấp']
    df_formatted[9] = df_crew['Ngày hết hạn']
    
    df_final = pd.concat([header_template, df_formatted], ignore_index=True)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
        
    return output.getvalue(), df_crew

st.set_page_config(page_title="Roster Automation", page_icon="✈️")
st.title("✈️ Roster Tool: Chuyển đổi file GD tự động")
st.markdown("Công cụ tự động chuyển đổi file General Declaration (GD) sang mẫu VNAPIS của Dịch vụ công quốc gia.")

TEMPLATE_PATH = "Template_VNAPIS.xlsx" 

uploaded_gd = st.file_uploader("Tải lên file GD (.xls hoặc .xlsx)", type=["xls", "xlsx"])

if uploaded_gd is not None:
    st.info("Đang xử lý dữ liệu...")
    try:
        excel_data, preview_data = process_roster_data(uploaded_gd, TEMPLATE_PATH)
        st.success("✅ Đã xử lý dữ liệu thành công! Vui lòng kiểm tra bản xem trước bên dưới:")
        st.dataframe(preview_data) 
        
        st.download_button(
            label="⬇️ Tải file Dịch vụ công đã điền (Excel)",
            data=excel_data,
            file_name="VNAPIS_Crew_AutoGenerated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Có lỗi xảy ra trong quá trình đọc file: {e}")
