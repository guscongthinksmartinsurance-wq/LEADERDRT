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
        # Thu nhieu kieu doc ngay thang khac nhau
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
        sheet_url = "https://docs.google.com/spreadsheets/d/1lbGUwZ7jd6dRZCLxJTgAQvPrezsCrAJs3ZcvTuneOwE"
        sh = client.open_by_url(sheet_url)

        ws1 = sh.get_worksheet_by_id(0)
        ws2 = sh.get_worksheet_by_id(680434099)
        ws3 = sh.get_worksheet_by_id(1751397007)

        df_mkt = pd.DataFrame(ws1.get_all_records())
        df_crm = pd.DataFrame(ws2.get_all_records())
        df_ml = pd.DataFrame(ws3.get_all_records())

        for df in [df_mkt, df_crm, df_ml]:
            if 'LEAD ID' in df.columns: df['MATCH_ID'] = df['LEAD ID'].apply(clean_id)
            if 'DATE ADDED' in df.columns: df['MY_REF'] = df['DATE ADDED'].apply(parse_month_year)

        # Lay danh sach thang tu tat ca cac sheet de anh chon
        all_months = set()
        for df in [df_mkt, df_crm, df_ml]:
            if 'MY_REF' in df.columns:
                all_months.update(df['MY_REF'].unique())
        
        # Loc bo cac gia tri loi khoi danh sach chon nhung van giu trong du lieu
        select_options = [m for m in all_months if m not in ["Khong co ngay", "Sai dinh dang", "Truoc 2024", "Loi doc ngay"]]
        sorted_months = sorted(select_options, key=lambda x: datetime.strptime(x, '%m/%Y'), reverse=True)
        
        # Them cac lua chon loi vao cuoi danh sach de anh kiem tra neu can
        sorted_months.extend([m for m in all_months if m in ["Khong co ngay", "Sai dinh dang", "Loi doc ngay"]])

        sel_month = st.sidebar.selectbox("Chon Thang/Nam", sorted_months)

        m_mkt = df_mkt[df_mkt['MY_REF'] == sel_month]
        m_crm = df_crm[df_crm['MY_REF'] == sel_month]
        m_ml = df_ml[df_ml['MY_REF'] == sel_month]

        tab1, tab2, tab3 = st.tabs(["Doi soat MKT", "Trang thai CRM", "Doanh thu Masterlife"])

        with tab1:
            st.subheader(f"Doi soat du lieu MKT - {sel_month}")
            if m_mkt.empty:
                st.warning(f"Sheet MKT khong co du lieu cho thang {sel_month}. Anh kiem tra cot DATE ADDED tai Sheet1.")
            else:
                total_mkt = len(m_mkt)
                leads_on_crm = m_mkt['MATCH_ID'].isin(df_crm['MATCH_ID']).sum()
                c1, c2 = st.columns(2)
                c1.metric("Tong Lead MKT vao", f"{total_mkt}")
                c2.metric("So luong da len CRM", f"{leads_on_crm}")
                st.dataframe(m_mkt[['OWNER', 'LEAD ID', 'CELLPHONE', 'DATE ADDED']], use_container_width=True)

        with tab2:
            st.subheader(f"Trang thai CRM - {sel_month}")
            if not m_crm.empty:
                pivot_status = m_crm.groupby(['OWNER', 'STATUS']).size().unstack(fill_value=0)
                st.dataframe(pivot_status.style.background_gradient(cmap="Blues", axis=1), use_container_width=True)
            else:
                st.info("Khong co du lieu CRM.")

        with tab3:
            st.subheader(f"Doanh thu Masterlife - {sel_month}")
            if m_ml.empty:
                st.warning(f"Sheet Masterlife khong co du lieu cho thang {sel_month}. Anh kiem tra cot DATE ADDED tai Sheet3.")
            else:
                m_ml['FINAL_REV'] = m_ml.apply(get_rev, axis=1)
                rev_funnel = m_ml[m_ml['SOURCE'].str.contains('Funnel', na=False, case=False)]['FINAL_REV'].sum()
                rev_cc = m_ml[m_ml['SOURCE'].str.contains('Cold Call|CC', na=False, case=False)]['FINAL_REV'].sum()
                
                c31, c32, c33 = st.columns(3)
                c31.metric("Doanh thu Funnel", f"${rev_funnel:,.0f}")
                c32.metric("Doanh thu Cold Call", f"${rev_cc:,.0f}")
                c33.metric("So ho so chot", f"{len(m_ml)}")
                st.dataframe(m_ml[['OWNER', 'LEAD ID', 'TEAM', 'FINAL_REV', 'DATE ADDED']], use_container_width=True)

    except Exception as e:
        st.error(f"Loi: {e}")

if __name__ == "__main__":
    main()
