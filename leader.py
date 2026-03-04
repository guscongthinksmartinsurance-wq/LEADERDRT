import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime
import io

def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_info = st.secrets["connections"]["gsheets"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(creds)

def clean_id(val):
    if pd.isna(val) or str(val).strip() == "": return ""
    s = str(val).strip().upper()
    if s.endswith('.0'): s = s[:-2]
    return s

def parse_month_year(date_val):
    """Hàm xử lý ngày tháng mạnh mẽ hơn"""
    try:
        if pd.isna(date_val) or str(date_val).strip() == "": return None
        # Thử ép kiểu ngày tháng, hỗ trợ nhiều định dạng (dayfirst=False cho giờ Mỹ)
        dt = pd.to_datetime(date_val, errors='coerce', dayfirst=False)
        if pd.notna(dt) and dt.year >= 2024: # Chỉ lấy từ năm 2024
            return dt.strftime('%m/%Y')
        return None
    except: return None

def get_rev(row):
    try:
        t_str = str(row.get('TARGET PREMIUM', 0)).replace(',', '').replace('$', '')
        a_str = str(row.get('ANNUAL PREMIUM', 0)).replace(',', '').replace('$', '')
        t_p = float(re.sub(r'[^0-9.]', '', t_str)) if t_str != "" else 0
        a_p = float(re.sub(r'[^0-9.]', '', a_str)) if a_str != "" else 0
        if t_p > 0 and a_p > 0: return min(t_p, a_p)
        return max(t_p, a_p)
    except: return 0

def main():
    st.set_page_config(page_title="TMC Strategic Portal", layout="wide")
    st.title("TMC Strategic Leader Dashboard")

    try:
        client = get_gspread_client()
        sheet_url = "https://docs.google.com/spreadsheets/d/1lbGUwZ7jd6dRZCLxJTgAQvPrezsCrAJs3ZcvTuneOwE"
        sh = client.open_by_url(sheet_url)

        ws1 = sh.get_worksheet_by_id(0)
        ws2 = sh.get_worksheet_by_id(680434099)
        ws3 = sh.get_worksheet_by_id(1751397007)

        df_mkt = pd.DataFrame(ws1.get_all_records())
        df_crm = pd.DataFrame(ws2.get_all_records())
        df_ml = pd.DataFrame(ws3.get_all_records())

        # Chuẩn hóa dữ liệu ngay khi nạp
        for df in [df_mkt, df_crm, df_ml]:
            if 'LEAD ID' in df.columns: df['MATCH_ID'] = df['LEAD ID'].apply(clean_id)
            if 'DATE ADDED' in df.columns: df['MY_REF'] = df['DATE ADDED'].apply(parse_month_year)

        # Lấy danh sách tháng từ 2024 trở đi
        all_months = set()
        for df in [df_mkt, df_crm, df_ml]:
            if 'MY_REF' in df.columns:
                all_months.update(df['MY_REF'].dropna().unique())
        
        # Sắp xếp và lọc Sidebar
        sorted_months = sorted(list(all_months), 
                               key=lambda x: datetime.strptime(x, '%m/%Y'), reverse=True)
        
        if not sorted_months:
            st.warning("Chưa tìm thấy dữ liệu hợp lệ từ năm 2024 trong cột DATE ADDED.")
            return

        sel_month = st.sidebar.selectbox("Chon Thang/Nam", sorted_months)

        # Lọc dữ liệu hiển thị
        m_mkt = df_mkt[df_mkt['MY_REF'] == sel_month]
        m_crm = df_crm[df_crm['MY_REF'] == sel_month]
        m_ml = df_ml[df_ml['MY_REF'] == sel_month]

        tab1, tab2, tab3 = st.tabs(["Doi soat MKT", "Trang thai CRM", "Doanh thu Masterlife"])

        with tab1:
            st.subheader(f"Doi soat du lieu MKT sang CRM - {sel_month}")
            total_mkt = len(m_mkt)
            # Đối soát LEAD ID (Sheet 1) vào TOÀN BỘ Sheet 2 (CRM) để không bị sót
            leads_on_crm = m_mkt['MATCH_ID'].isin(df_crm['MATCH_ID']).sum()
            
            c1, c2 = st.columns(2)
            c1.metric("Tong Lead MKT vao", f"{total_mkt}")
            c2.metric("So luong da len CRM", f"{leads_on_crm}", f"{leads_on_crm/total_mkt*100:.1f}%" if total_mkt > 0 else "0%")
            st.dataframe(m_mkt[['OWNER', 'LEAD ID', 'CELLPHONE', 'DATE ADDED']], use_container_width=True)

        with tab2:
            st.subheader(f"Thong ke Status va Owner - {sel_month}")
            if not m_crm.empty:
                pivot_status = m_crm.groupby(['OWNER', 'STATUS']).size().unstack(fill_value=0)
                st.dataframe(pivot_status.style.background_gradient(cmap="Blues", axis=1), use_container_width=True)
            else:
                st.info("Khong co du lieu CRM trong thang nay.")

        with tab3:
            st.subheader(f"Doanh thu va Toc do chot - {sel_month}")
            if not m_ml.empty:
                m_ml['FINAL_REV'] = m_ml.apply(get_rev, axis=1)
                
                def get_cycle(row):
                    try:
                        crm_record = df_crm[df_crm['MATCH_ID'] == row['MATCH_ID']]
                        if not crm_record.empty:
                            # Ưu tiên lấy ngày đầu tiên xuất hiện trên CRM
                            d_crm = pd.to_datetime(crm_record['DATE ADDED'].iloc[0], errors='coerce')
                            d_ml = pd.to_datetime(row['DATE ADDED'], errors='coerce')
                            if pd.notna(d_crm) and pd.notna(d_ml):
                                return (d_ml.year - d_crm.year) * 12 + (d_ml.month - d_crm.month)
                        return "N/A"
                    except: return "N/A"

                m_ml['CYCLE'] = m_ml.apply(get_cycle, axis=1)
                
                rev_funnel = m_ml[m_ml['SOURCE'].str.contains('Funnel', na=False, case=False)]['FINAL_REV'].sum()
                rev_cc = m_ml[m_ml['SOURCE'].str.contains('Cold Call|CC', na=False, case=False)]['FINAL_REV'].sum()

                c31, c32, c33 = st.columns(3)
                c31.metric("Doanh thu Funnel", f"${rev_funnel:,.0f}")
                c32.metric("Doanh thu Cold Call", f"${rev_cc:,.0f}")
                c33.metric("So ho so chot", f"{len(m_ml)}")
                st.dataframe(m_ml[['OWNER', 'LEAD ID', 'TEAM', 'FINAL_REV', 'CYCLE']], use_container_width=True)

        if st.sidebar.button("Xuat Bao Cao Excel"):
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                if not m_ml.empty:
                    export_df = m_ml[['OWNER', 'LEAD ID', 'SOURCE', 'TEAM', 'FINAL_REV', 'CYCLE']]
                    export_df.to_excel(writer, sheet_name='TMC_Report', index=False, startrow=3)
                    workbook = writer.book
                    worksheet = writer.sheets['TMC_Report']
                    title_fmt = workbook.add_format({'bold': True, 'font_size': 14})
                    header_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})
                    worksheet.write(0, 0, f"BAO CAO TMC - {sel_month}", title_fmt)
                    worksheet.write(1, 0, f"Ngay xuat: {datetime.now().strftime('%d/%m/%Y')}")
                    for col_num, value in enumerate(export_df.columns.values):
                        worksheet.write(3, col_num, value, header_fmt)
                        worksheet.set_column(col_num, col_num, 18)
            st.sidebar.download_button("Tai File", buf.getvalue(), f"TMC_Report_{sel_month.replace('/','_')}.xlsx")

    except Exception as e:
        st.error(f"Loi: {e}")

if __name__ == "__main__":
    main()
