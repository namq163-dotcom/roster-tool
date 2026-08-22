import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO, StringIO
from openpyxl import load_workbook

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
    content = gd_file.getvalue()
    df_gd = None
    
    # 1. Đọc file GD (Hỗ trợ Excel, HTML, Text)
    try:
        df_gd = pd.read_excel(BytesIO(content))
    except Exception:
        pass

    if df_gd is None or len(df_gd) == 0:
        for enc in ['utf-8', 'latin1', 'cp1258', 'utf-16']:
            try:
                html_content = content.decode(enc, errors='ignore')
                dfs = pd.read_html(StringIO(html_content))
                if len(dfs) > 0:
                    df_gd = dfs[0]
                    break
            except:
                continue

    if df_gd is None or len(df_gd) == 0:
        for sep in ['\t', ',', ';', '|']:
            try:
                df_gd = pd.read_csv(BytesIO(content), sep=sep, encoding='latin1', on_bad_lines='skip')
                if df_gd is not None and len(df_gd.columns) > 1:
                    break
            except:
                continue

    if df_gd is None or len(df_gd) == 0:
        raise ValueError("Không thể đọc được dữ liệu trong file GD.")

    # 2. Tìm bảng chứa danh sách tổ bay
    header_idx = None
    for idx, row in df_gd.iterrows():
        row_str = row.fillna("").astype(str).str.lower()
        if any('passport' in s for s in row_str) and any('name' in s for s in row_str):
            header_idx = idx
            break
            
    if header_idx is None:
        raise ValueError("Không tìm thấy bảng danh sách tổ bay trong file GD.")
        
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
        if 'declaration of health' in row.fillna("").astype(str).str.cat(sep=" ").lower():
            break
            
        name_val = str(row.iloc[col_name]).strip() if col_name is not None and pd.notna(row.iloc[col_name]) else 'nan'
        passport_val = str(row.iloc[col_passport]).strip() if col_passport is not None and pd.notna(row.iloc[col_passport]) else 'nan'
        
        if name_val.lower() != 'nan' and passport_val.lower() != 'nan' and name_val != '':
            name_parts = name_val.split()
            if len(name_parts) > 1 and len(name_parts[-1]) <= 3 and name_parts[-1].isupper():
                name_parts = name_parts[:-1]
                
            family_name = name_parts[0] if name_parts else ""
            given_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
            
            nat = str(row.iloc[col_nat]).strip() if col_nat is not None and pd.notna(row.iloc[col_nat]) else ""
            if nat.upper() == 'VNM': nat = 'VN'
            
            gender = str(row.iloc[col_gender]).strip() if col_gender is not None and pd.notna(row.iloc[col_gender]) else ""
            dob = parse_date(row.iloc[col_dob]) if col_dob is not None else ""
            expiry = parse_date(row.iloc[col_expiry]) if col_expiry is not None else ""
            
            crew_data.append([
                family_name, 
                None, 
                given_name, 
                gender, 
                nat, 
                dob, 
                'P', 
                passport_val, 
                nat, 
                expiry
            ])

    if len(crew_data) == 0:
        raise ValueError("Không trích xuất được dữ liệu tổ bay nào từ file GD.")

    # 3. Ghi trực tiếp vào file Template giữ nguyên 100% định dạng gốc
    output = BytesIO()
    book = load_workbook(template_file_path)
    sheet = book.active
    
    # Bắt đầu điền dữ liệu từ dòng 13 (Cột A đến J tương ứng cột 1 đến 10)
    start_row = 13 
    for r_idx, row_data in enumerate(crew_data, start_row):
        for c_idx, value in enumerate(row_data, 1):
            sheet.cell(row=r_idx, column=c_idx, value=value)
            
    book.save(output)
    
    # Trả về dữ liệu kèm bảng preview
    df_preview = pd.DataFrame(crew_data, columns=['Họ', 'Tên đệm', 'Tên', 'Giới tính', 'Quốc tịch', 'Ngày sinh', 'Loại giấy tờ', 'Số giấy tờ', 'Nơi cấp', 'Ngày hết hạn'])
    return output.getvalue(), df_preview

st.set_page_config(page_title="Roster Automation", page_icon="✈️")
st.title("✈️ Roster Tool: Chuyển đổi file GD tự động")
st.markdown("Công cụ tự động chuyển đổi file General Declaration (GD) sang mẫu VNAPIS của Dịch vụ công quốc gia.")

TEMPLATE_PATH = "Template_VNAPIS.xlsx" 

uploaded_gd = st.file_uploader("Tải lên file GD (.xls, .xlsx)", type=["xls", "xlsx", "txt", "csv"])

if uploaded_gd is not None:
    st.info("Đang xử lý dữ liệu...")
    try:
        excel_data, preview_data = process_roster_data(uploaded_gd, TEMPLATE_PATH)
        st.success("✅ Đã xử lý dữ liệu thành công! Vui lòng kiểm tra bản xem trước bên dưới:")
        st.dataframe(preview_data) 
        
        st.download_button(
            label="⬇️ Tải file Dịch vụ công chuẩn mẫu (Excel)",
            data=excel_data,
            file_name="VNAPIS_Crew_Completed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Có lỗi xảy ra trong quá trình đọc file: {e}")
