import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime
import io

# --- 1. CẤU HÌNH GIAO DIỆN & STYLE (MIDNIGHT BLUE THEME) ---
st.set_page_config(page_title="TMC Strategic Portal", layout="wide")

st.markdown("""
    <style>
    /* Tổng thể và Sidebar */
    .stApp { background-color: #F8F9FA; }
    section[data-testid="stSidebar"] { background-color: #112233 !important; }
    section[data-testid="stSidebar"] .stMarkdown, section[data-testid="stSidebar"] label { color: white !important; }
    
    /* Chỉnh sửa Metric (Thẻ chỉ số) */
    div[data-testid="stMetric"] {
        background-color: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        border-left: 5px solid #1F4E78;
    }
    div[data-testid="stMetricLabel"] { font-size: 16px !important; color: #555 !important; font-weight: bold; }
    div[data-testid="stMetricValue"] { color: #1F4E78 !important; }

    /* Tiêu đề và Tabs */
    h1 { color: #112233; font-family: 'Georgia', serif; font-weight: bold; }
    h3 { color: #1F4E78; border-bottom: 2px solid #D9E1F2; padding-bottom: 10px; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #E9ECEF;
        border-radius: 5px 5px 0 0;
        padding: 10px 20px;
        color: #495057;
    }
    .stTabs [aria-selected="true"] { background-color: #1F4E78 !important; color: white !important; }
    
    /* Chỉnh sửa bảng DataFrame */
    .stDataFrame { border: 1px solid #E9ECEF; border-radius: 10px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. LOGIC XỬ LÝ DỮ LIỆU (GIỮ NGUYÊN) ---
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
    st.title("TMC Strategic Leader Dashboard")
    
    # Hiển thị thời gian thực theo phong cách sâu sắc
    now_cst = datetime.now().strftime('%d/%m/%Y | %H:%M')
    st.sidebar.caption(f"Đồng bộ dữ liệu: {now_cst}")

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

        sel_month = st.sidebar.selectbox("Chọn Tháng/Năm quản trị", sorted_months)

        m_mkt = df_mkt[df_mkt['MY_REF'] == sel_month]
        m_crm = df_crm[df_crm['MY_REF'] == sel_month]
        m_ml = df_ml[df_ml['MY_REF'] == sel_month]

        tab1, tab2, tab3 = st.tabs(["🎯 ĐỐI SOÁT MARKETING", "📊 CRM DASHBOARD", "💰 SALES PERFORMANCE"])

        with tab1:
            st.subheader(f"Kiểm toán nguồn Marketing - {sel_month}")
            if not m_mkt.empty:
                mask_on_crm = m_mkt['MATCH_ID'].isin(df_crm['MATCH_ID'])
                leads_on_crm_count = mask_on_crm.sum()
                df_missing = m_mkt[~mask_on_crm].copy()
                
                c1, c2 = st.columns(2)
                c1.metric("Tổng Lead Marketing", f"{len(m_mkt)}")
                c2.metric("Đã cập nhật CRM", f"{leads_on_crm_count}", f"{leads_on_crm_count/len(m_mkt)*100:.1f}%")
                
                if not df_missing.empty:
                    st.markdown(f"<div style='background-color:#FDECEA; padding:15px; border-radius:10px; border-left:5px solid #990000; color:#990000; font-weight:bold;'>⚠️ DANH SÁCH {len(df_missing)} LEAD CHƯA LÊN CRM</div>", unsafe_allow_html=True)
                    df_missing.insert(0, 'STT', range(1, len(df_missing) + 1))
                    st.dataframe(df_missing[['STT', 'OWNER', 'LEAD ID', 'CELLPHONE', 'DATE ADDED']], use_container_width=True, hide_index=True)
                
                st.markdown("#### Toàn bộ danh sách Marketing")
                display_mkt = m_mkt.copy()
                display_mkt.insert(0, 'STT', range(1, len(display_mkt) + 1))
                st.dataframe(display_mkt[['STT', 'OWNER', 'LEAD ID', 'CELLPHONE', 'DATE ADDED']], use_container_width=True, hide_index=True)

        with tab2:
            st.subheader(f"Phân tích Pipeline & Trạng thái - {sel_month}")
            if not m_crm.empty:
                col_chart1, col_chart2 = st.columns([1, 1])
                with col_chart1:
                    st.write("**Trọng số Trạng thái (Status)**")
                    st.bar_chart(m_crm['STATUS'].value_counts(), color="#1F4E78")
                with col_chart2:
                    st.write("**Mật độ Lead theo Owner**")
                    st.bar_chart(m_crm['OWNER'].value_counts(), color="#2E7D32")
                
                pivot_status = m_crm.groupby(['OWNER', 'STATUS']).size().unstack(fill_value=0)
                st.dataframe(pivot_status.style.background_gradient(cmap="Blues", axis=1), use_container_width=True)

        with tab3:
            st.subheader(f"Kết quả kinh doanh Masterlife - {sel_month}")
            if not m_ml.empty:
                m_ml['FINAL_REV'] = m_ml.apply(get_rev, axis=1)
                
                def get_cycle_new(row):
                    source_val = str(row.get('SOURCE', '')).upper()
                    if 'COLD CALL' in source_val or 'CC' in source_val: return "CC"
                    crm_rec = df_crm[df_crm['MATCH_ID'] == row['MATCH_ID']]
                    if not crm_rec.empty:
                        d_crm = pd.to_datetime(crm_rec['DATE ADDED'].iloc[0], errors='coerce')
                        d_ml = pd.to_datetime(row['DATE ADDED'], errors='coerce')
                        if pd.notna(d_crm) and pd.notna(d_ml):
                            diff = (d_ml.year - d_crm.year) * 12 + (d_ml.month - d_crm.month)
                            return diff if diff >= 0 else 0
                    return "N/A"

                m_ml['CYCLE'] = m_ml.apply(get_cycle_new, axis=1)
                
                # Metric Doanh thu thẻ lớn
                rev_total = m_ml['FINAL_REV'].sum()
                st.metric("TỔNG DOANH THU THÁNG", f"${rev_total:,.0f}")

                display_ml = m_ml[['OWNER', 'LEAD ID', 'TEAM', 'FINAL_REV', 'CYCLE']].copy()
                display_ml.insert(0, 'STT', range(1, len(display_ml) + 1))
                st.dataframe(display_ml, use_container_width=True, hide_index=True)

        st.sidebar.markdown("---")
        if st.sidebar.button("📦 XUẤT BÁO CÁO CHIẾN LƯỢC"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                title_fmt = workbook.add_format({'bold': True, 'font_size': 14})

                def write_sheet(df, name, title):
                    df_export = df.copy()
                    if 'STT' not in df_export.columns and name != "CRM_Status":
                        df_export.insert(0, 'STT', range(1, len(df_export) + 1))
                    df_export.to_excel(writer, sheet_name=name, index=True if name == "CRM_Status" else False, startrow=3)
                    ws = writer.sheets[name]
                    ws.write(0, 0, title, title_fmt)
                    ws.write(1, 0, f"Ngay xuat: {datetime.now().strftime('%d/%m/%Y')}")
                    for col_num, value in enumerate(df_export.columns.values):
                        col_idx = col_num + (1 if name == "CRM_Status" else 0)
                        ws.write(3, col_idx, value, header_fmt)
                        ws.set_column(col_idx, col_idx, 20, cell_fmt)
                    if name == "CRM_Status": ws.write(3, 0, "OWNER", header_fmt)

                if not m_ml.empty: write_sheet(m_ml[['OWNER', 'LEAD ID', 'SOURCE', 'TEAM', 'FINAL_REV', 'CYCLE']], 'Sales_Summary', f"BAO CAO DOANH THU - {sel_month}")
                if not m_crm.empty: write_sheet(pivot_status, 'CRM_Status', f"THONG KE TRANG THAI CRM - {sel_month}")
                if not m_mkt.empty: write_sheet(m_mkt[['OWNER', 'LEAD ID', 'CELLPHONE', 'DATE ADDED']], 'MKT_Audit', f"DOI SOAT LEAD MKT - {sel_month}")

            st.sidebar.download_button("📥 Tải báo cáo đa tầng", output.getvalue(), f"TMC_Report_{sel_month.replace('/','_')}.xlsx")

    except Exception as e:
        st.error(f"Hệ thống tạm thời gián đoạn: {e}")

if __name__ == "__main__":
    main()
