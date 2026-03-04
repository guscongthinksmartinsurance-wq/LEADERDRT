import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime
import io

# --- 1. KẾT NỐI GOOGLE SHEETS BẢO MẬT ---
def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # Đọc thông tin từ Streamlit Secrets
    creds_info = st.secrets["connections"]["gsheets"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(creds)

# --- 2. HÀM LÀM SẠCH VÀ XỬ LÝ DỮ LIỆU ---
def clean_id(val):
    if pd.isna(val) or str(val).strip() == "": return ""
    s = str(val).strip().upper()
    if s.endswith('.0'): s = s[:-2]
    return s

def parse_month_year(date_val):
    try:
        if pd.isna(date_val) or str(date_val).strip() == "": return None
        # Xử lý định dạng Mỹ MM/DD/YYYY từ Sheet
        dt = pd.to_datetime(date_val, errors='coerce')
        return dt.strftime('%m/%Y') if pd.notna(dt) else None
    except: return None

# --- 3. ENGINE PHÂN TÍCH CHÍNH ---
def main():
    st.title("🚀 TMC Strategic Leader Dashboard")
    
    try:
        client = get_gspread_client()
        sheet_url = "https://docs.google.com/spreadsheets/d/1lbGUwZ7jd6dRZCLxJTgAQvPrezsCrAJs3ZcvTuneOwE"
        sh = client.open_by_url(sheet_url)

        # Đọc 3 Worksheets bằng GID
        ws1 = sh.get_worksheet_by_id(0)          # Sheet1: MKT
        ws2 = sh.get_worksheet_by_id(680434099)  # Sheet2: CRM
        ws3 = sh.get_worksheet_by_id(1751397007) # Sheet3: MASTERLIFE

        df_mkt = pd.DataFrame(ws1.get_all_records())
        df_crm = pd.DataFrame(ws2.get_all_records())
        df_ml = pd.DataFrame(ws3.get_all_records())

        # Tiền xử lý dữ liệu an toàn
        for df in [df_mkt, df_crm, df_ml]:
            if 'LEAD ID' in df.columns: df['MATCH_ID'] = df['LEAD ID'].apply(clean_id)
            if 'DATE ADDED' in df.columns: df['MY'] = df['DATE ADDED'].apply(parse_month_year)

        # Lấy danh sách Tháng/Năm để lọc
        available_months = sorted([m for m in df_crm['MY'].unique() if m is not None], 
                                  key=lambda x: datetime.strptime(x, '%m/%Y'), reverse=True)
        
        sel_month = st.sidebar.selectbox("📅 Chọn Tháng/Năm", available_months)
        
        # Filter dữ liệu theo tháng được chọn
        m_mkt = df_mkt[df_mkt['MY'] == sel_month]
        m_crm = df_crm[df_crm['MY'] == sel_month]
        m_ml = df_ml[df_ml['MY'] == sel_month]

        # --- HIỂN THỊ BÁO CÁO ---
        t1, t2, t3 = st.tabs(["🎯 Đối soát MKT", "🏢 Trạng thái CRM", "💰 Doanh thu Masterlife"])

        with t1:
            st.subheader(f"Đối soát dữ liệu MKT sang CRM - {sel_month}")
            total_mkt = len(m_mkt)
            # Kiểm tra LEAD ID của MKT (Sheet 1) đã xuất hiện trong CRM (toàn bộ Sheet 2) chưa
            match_count = m_mkt['MATCH_ID'].isin(df_crm['MATCH_ID']).sum()
            
            c1, c2 = st.columns(2)
            c1.metric("Tổng Lead MKT vào", f"{total_mkt} Lead")
            c2.metric("Số lượng đã lên CRM", f"{match_count} Lead", f"{match_count/total_mkt*100:.1f}%" if total_mkt > 0 else "0%")

        with t2:
            st.subheader("Thống kê Trạng thái & Nhân viên (Owner)")
            if not m_crm.empty:
                # Liệt kê tất cả Status không lược bớt
                pivot_status = m_crm.groupby(['OWNER', 'STATUS']).size().unstack(fill_value=0)
                st.dataframe(pivot_status.style.background_gradient(axis=1, cmap="Blues"), use_container_width=True)
            else:
                st.info("Tháng này chưa có dữ liệu trên CRM.")

        with t3:
            st.subheader("Hiệu suất Doanh thu & Tốc độ chốt")
            
            def process_ml_row(row):
                # Logic Doanh thu: lấy min của Target và Annual
                try:
                    t_p = float(str(row.get('TARGET PREMIUM', 0)).replace(',', ''))
                    a_p = float(str(row.get('ANNUAL PREMIUM', 0)).replace(',', ''))
                    rev = min(t_p, a_p) if t_p > 0 and a_p > 0 else max(t_p, a_p)
                    
                    # Tính độ trễ (Cycle)
                    cycle = "N/A"
                    crm_entry = df_crm[df_crm['MATCH_ID'] == row['MATCH_ID']]
                    if not crm_entry.empty:
                        d_crm = pd.to_datetime(crm_entry['DATE ADDED'].iloc[0])
                        d_ml = pd.to_datetime(row['DATE ADDED'])
                        cycle = (d_ml.year - d_crm.year) * 12 + (d_ml.month - d_crm.month)
                    
                    return pd.Series([rev, cycle])
                except: return pd.Series([0, "N/A"])

            if not m_ml.empty:
                m_ml[['REV_FINAL', 'CYCLE']] = m_ml.apply(process_ml_row, axis=1)
                
                # Phân loại nguồn
                rev_funnel = m_ml[m_ml['SOURCE'].str.contains('Funnel', na=False, case=False)]['REV_FINAL'].sum()
                rev_cc = m_ml[m_ml['SOURCE'].str.contains('Cold Call|CC', na=False, case=False)]['REV_FINAL'].sum()
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Doanh thu Funnel", f"${rev_funnel:,.0f}")
                c2.metric("Doanh thu Cold Call", f"${rev_cc:,.0f}")
                c3.metric("Tổng Case chốt", len(m_ml))
                
                st.markdown("### Chi tiết Owner & Tốc độ chốt")
                st.dataframe(m_ml[['OWNER', 'LEAD ID', 'TEAM', 'REV_FINAL', 'CYCLE']].rename(columns={'CYCLE': 'Trễ (Tháng)'}), use_container_width=True)

        # --- NÚT EXPORT FILE ---
        st.sidebar.divider()
        if st.sidebar.button("📊 Xuất Báo Cáo Excel"):
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                # Tạo nội dung export
                export_df = m_ml[['OWNER', 'LEAD ID', 'SOURCE', 'TEAM', 'REV_FINAL', 'CYCLE']] if not m_ml.empty else pd.DataFrame()
                export_df.to_excel(writer, sheet_name='Sales_Report', index=False, startrow=3)
                
                workbook = writer.book
                ws = writer.sheets['Sales_Report']
                
                # Định dạng cổ điển & chuyên nghiệp
                title_fmt = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#1F4E78'})
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
                
                ws.write(0, 0, f"TMC STRATEGIC REPORT - {sel_month}", title_fmt)
                ws.write(1, 0, f"Export Date: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
                
                for col_num, value in enumerate(export_df.columns.values):
                    ws.write(3, col_num, value, header_fmt)
                    ws.set_column(col_num, col_num, 18) # Khoảng cách cột đều nhau

            st.sidebar.download_button("📥 Tải File Ngay", buf.getvalue(), f"TMC_Report_{sel_month.replace('/','_')}.xlsx")

    except Exception as e:
        st.error(f"Lỗi kết nối hoặc dữ liệu: {e}")
        st.info("Anh kiểm tra lại quyền chia sẻ Sheet cho Email Service Account nhé!")

if __name__ == "__main__":
    main()