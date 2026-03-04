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
    try:
        if pd.isna(date_val) or str(date_val).strip() == "": return "Khong co ngay"
        dt = pd.to_datetime(date_val, errors='coerce')
        if pd.notna(dt):
            if dt.year < 2024: return "Truoc 2024"
            return dt.strftime('%m/%Y')
        return "Sai dinh dang"
    except: return "Loi doc ngay"

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
        sh = client.open_by_url("https://docs.google.com/spreadsheets/d/1lbGUwZ7jd6dRZCLxJTgAQvPrezsCrAJs3ZcvTuneOwE")

        df_mkt = pd.DataFrame(sh.get_worksheet_by_id(0).get_all_records())
        df_crm = pd.DataFrame(sh.get_worksheet_by_id(680434099).get_all_records())
        df_ml = pd.DataFrame(sh.get_worksheet_by_id(1751397007).get_all_records())

        for df in [df_mkt, df_crm, df_ml]:
            if 'LEAD ID' in df.columns: df['MATCH_ID'] = df['LEAD ID'].apply(clean_id)
            if 'DATE ADDED' in df.columns: df['MY_REF'] = df['DATE ADDED'].apply(parse_month_year)

        all_months = set()
        for df in [df_mkt, df_crm, df_ml]:
            if 'MY_REF' in df.columns: all_months.update(df['MY_REF'].unique())
        
        valid_options = [m for m in all_months if m not in ["Khong co ngay", "Sai dinh dang", "Truoc 2024", "Loi doc ngay"]]
        sorted_months = sorted(valid_options, key=lambda x: datetime.strptime(x, '%m/%Y'), reverse=True)
        sorted_months.extend([m for m in all_months if m in ["Khong co ngay", "Sai dinh dang", "Loi doc ngay"]])

        sel_month = st.sidebar.selectbox("Chon Thang/Nam", sorted_months)

        m_mkt = df_mkt[df_mkt['MY_REF'] == sel_month]
        m_crm = df_crm[df_crm['MY_REF'] == sel_month]
        m_ml = df_ml[df_ml['MY_REF'] == sel_month]

        tab1, tab2, tab3 = st.tabs(["Doi soat MKT", "Trang thai CRM", "Doanh thu Masterlife"])

        # LOGIC TINH TOAN MASTERLIFE
        if not m_ml.empty:
            m_ml['FINAL_REV'] = m_ml.apply(get_rev, axis=1)
            def get_cycle(row):
                crm_rec = df_crm[df_crm['MATCH_ID'] == row['MATCH_ID']]
                if not crm_rec.empty:
                    d_crm = pd.to_datetime(crm_rec['DATE ADDED'].iloc[0], errors='coerce')
                    d_ml = pd.to_datetime(row['DATE ADDED'], errors='coerce')
                    if pd.notna(d_crm) and pd.notna(d_ml): return (d_ml.year - d_crm.year) * 12 + (d_ml.month - d_crm.month)
                return "N/A"
            m_ml['CYCLE'] = m_ml.apply(get_cycle, axis=1)

        with tab1:
            st.subheader(f"Doi soat MKT - {sel_month}")
            if not m_mkt.empty:
                leads_on_crm = m_mkt['MATCH_ID'].isin(df_crm['MATCH_ID']).sum()
                st.metric("Lead da len CRM", f"{leads_on_crm}/{len(m_mkt)}")
                st.dataframe(m_mkt[['OWNER', 'LEAD ID', 'CELLPHONE', 'DATE ADDED']], use_container_width=True)

        with tab2:
            st.subheader(f"Trang thai CRM - {sel_month}")
            pivot_status = m_crm.groupby(['OWNER', 'STATUS']).size().unstack(fill_value=0) if not m_crm.empty else pd.DataFrame()
            st.dataframe(pivot_status, use_container_width=True)

        with tab3:
            st.subheader(f"Doanh thu - {sel_month}")
            if not m_ml.empty:
                st.dataframe(m_ml[['OWNER', 'LEAD ID', 'TEAM', 'FINAL_REV', 'CYCLE']], use_container_width=True)

        # --- EXPORT MULTI-SHEET ---
        st.sidebar.markdown("---")
        if st.sidebar.button("Chuan bi File Export"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # Dinh dang
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                title_fmt = workbook.add_format({'bold': True, 'font_size': 14})

                def write_sheet(df, name, title):
                    df.to_excel(writer, sheet_name=name, index=True if name == "CRM_Status" else False, startrow=3)
                    ws = writer.sheets[name]
                    ws.write(0, 0, title, title_fmt)
                    ws.write(1, 0, f"Ngay xuat: {datetime.now().strftime('%d/%m/%Y')}")
                    for col_num, value in enumerate(df.columns.values):
                        col_idx = col_num + (1 if name == "CRM_Status" else 0)
                        ws.write(3, col_idx, value, header_fmt)
                        ws.set_column(col_idx, col_idx, 20, cell_fmt)
                    if name == "CRM_Status": ws.write(3, 0, "OWNER", header_fmt)

                # Sheet 1: Masterlife
                if not m_ml.empty: write_sheet(m_ml[['OWNER', 'LEAD ID', 'SOURCE', 'TEAM', 'FINAL_REV', 'CYCLE']], 'Sales_Summary', f"BAO CAO DOANH THU - {sel_month}")
                # Sheet 2: CRM Status
                if not pivot_status.empty: write_sheet(pivot_status, 'CRM_Status', f"THONG KE TRANG THAI CRM - {sel_month}")
                # Sheet 3: MKT Audit
                if not m_mkt.empty: write_sheet(m_mkt[['OWNER', 'LEAD ID', 'CELLPHONE', 'DATE ADDED']], 'MKT_Audit', f"DOI SOAT LEAD MKT - {sel_month}")

            st.sidebar.download_button("Tai File Excel (Da Sheet)", output.getvalue(), f"TMC_Full_Report_{sel_month.replace('/','_')}.xlsx")

    except Exception as e:
        st.error(f"Loi: {e}")

if __name__ == "__main__":
    main()
