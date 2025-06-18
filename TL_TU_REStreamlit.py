import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus
from datetime import datetime, timedelta
from io import BytesIO
import xlsxwriter  # for formatting (used implicitly via ExcelWriter)

# Database Configuration
CONFIG = {
    'server': 'caappsdb,1435',
    'database': 'BCSSoft_ConAppSys',
    'username': 'Deepak',
    'password': 'Deepak@321',
    'driver': 'ODBC Driver 17 for SQL Server'
}

# Establishing Database Connection
def get_db_engine():
    connection_string = (
        f"mssql+pyodbc://{CONFIG['username']}:%s@{CONFIG['server']}/{CONFIG['database']}?"
        f"driver={quote_plus(CONFIG['driver'])}&TrustServerCertificate=yes"
    ) % quote_plus(CONFIG['password'])
    return create_engine(connection_string)

# Fetch ORP Data
def fetch_orp_data(engine):
    today = datetime.today()
    three_months_ago = today - timedelta(days=90)

    query = text("""
        SELECT
            rh.CounterCode,
            rh.ShipFromCode,
            th.TONo,
            rh.ReceiptNo,
            th.LoadClosedDt, 
            th.UnloadClosedDt,
            rh.ClosedDt AS ClosedReceipt
        FROM 
            [BCSSoft_ConAppSys].[dbo].[tbTruckLoadHeader] th
        LEFT JOIN 
            [BCSSoft_ConAppSys].[dbo].[tbReceiptHeader] rh 
            ON th.TONo = rh.ReceiptNo
        WHERE 
            th.SourceFrom = 'orp' 
            AND th.LoadClosedDt >= :start_date
            AND th.LoadClosedDt <= :end_date
    """)

    with engine.connect() as conn:
        df = pd.read_sql(query, conn, params={'start_date': three_months_ago, 'end_date': today})

    df['Unload Status'] = df['UnloadClosedDt'].apply(lambda x: 'New' if pd.isna(x) else 'Closed')
    df['Receipt Status'] = df['ClosedReceipt'].apply(lambda x: 'New' if pd.isna(x) else 'Closed')

    # Duration columns
    df['Load vs Unload Duration (Days)'] = df.apply(
        lambda row: (row['UnloadClosedDt'] - row['LoadClosedDt']).days if pd.notna(row['UnloadClosedDt']) and pd.notna(row['LoadClosedDt']) else None, axis=1)
    df['Load vs Unload Duration (Hours)'] = df.apply(
        lambda row: (row['UnloadClosedDt'] - row['LoadClosedDt']).seconds // 3600 if pd.notna(row['UnloadClosedDt']) and pd.notna(row['LoadClosedDt']) else None, axis=1)
    df['Load vs Unload Duration (Minutes)'] = df.apply(
        lambda row: ((row['UnloadClosedDt'] - row['LoadClosedDt']).seconds % 3600) // 60 if pd.notna(row['UnloadClosedDt']) and pd.notna(row['LoadClosedDt']) else None, axis=1)

    df['Unload vs Receipt (Days)'] = df.apply(
        lambda row: (row['ClosedReceipt'] - row['UnloadClosedDt']).days if pd.notna(row['ClosedReceipt']) and pd.notna(row['UnloadClosedDt']) else None, axis=1)
    df['Unload vs Receipt (Hours)'] = df.apply(
        lambda row: (row['ClosedReceipt'] - row['UnloadClosedDt']).seconds // 3600 if pd.notna(row['ClosedReceipt']) and pd.notna(row['UnloadClosedDt']) else None, axis=1)
    df['Unload vs Receipt (Minutes)'] = df.apply(
        lambda row: ((row['ClosedReceipt'] - row['UnloadClosedDt']).seconds % 3600) // 60 if pd.notna(row['ClosedReceipt']) and pd.notna(row['UnloadClosedDt']) else None, axis=1)

    # Reorder columns
    column_order = [
        'CounterCode', 'ShipFromCode', 'TONo', 'ReceiptNo', 'LoadClosedDt', 'UnloadClosedDt',
        'ClosedReceipt', 'Unload Status', 'Receipt Status',
        'Load vs Unload Duration (Days)', 'Load vs Unload Duration (Hours)', 'Load vs Unload Duration (Minutes)',
        'Unload vs Receipt (Days)', 'Unload vs Receipt (Hours)', 'Unload vs Receipt (Minutes)'
    ]
    return df[column_order]

# Excel Export Function
def to_excel(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='ORP Data', index=False, startrow=2)
        workbook = writer.book
        worksheet = writer.sheets['ORP Data']

        # Column grouping headers
        load_unload_start = df.columns.get_loc('Load vs Unload Duration (Days)')
        load_unload_end = df.columns.get_loc('Load vs Unload Duration (Minutes)')
        receipt_start = df.columns.get_loc('Unload vs Receipt (Days)')
        receipt_end = df.columns.get_loc('Unload vs Receipt (Minutes)')

        merged_format = workbook.add_format({
            'bold': True, 'font_color': 'white', 'fg_color': '#4472C4',
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
        })

        worksheet.merge_range(1, load_unload_start, 1, load_unload_end, "Duration to Complete TruckUnload", merged_format)
        worksheet.merge_range(1, receipt_start, 1, receipt_end, "Duration to Complete Receipt", merged_format)

        # Write column headers (row 2)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(2, col_num, value, header_format)

        # Adjust column widths
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
            worksheet.set_column(idx, idx, min(max_len, 30))

    output.seek(0)
    return output

# Streamlit App
def main():
    st.set_page_config(page_title="ORP Delay Report", layout="wide")
    st.title("🚚 ORP Delay Report - Last 3 Months")

    try:
        with st.spinner("Connecting to database..."):
            engine = get_db_engine()

        with st.spinner("Fetching data..."):
            df = fetch_orp_data(engine)

        if df.empty:
            st.warning("No ORP data found in the last 3 months.")
            return

        st.success(f"Fetched {len(df)} records.")
        st.dataframe(df, use_container_width=True)

        # Excel download
        excel_file = to_excel(df)
        st.download_button(
            label="📥 Download ORP Report (Excel)",
            data=excel_file,
            file_name="ORP_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")

if __name__ == '__main__':
    main()
