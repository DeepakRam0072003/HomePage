import streamlit as st
import pandas as pd
import pyodbc
from io import BytesIO

# DB connection config
server = 'caappsdb,1435'
database = 'BCSSoft_ConAppSys'
username = 'Deepak'
password = 'Deepak@321'

def get_db_connection():
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};DATABASE={database};UID={username};PWD={password}"
    )
    conn = pyodbc.connect(conn_str)
    return conn

def run_query(conn):
    query = """
    SELECT 
        tlh.TruckLoadStsCode,
        tlh.CounterCode AS [Transfer-From],
        toh.ShipToCode,
        tlh.TruckLoadHeaderId,
        tlh.TONo,
        tlh.HostHeaderNo,
        tlh.SourceFrom,
        tlh.CreatedDt,
        tlh.NavTONo,
        CASE
            WHEN tlh.NavTONo LIKE 'EDTO%' THEN 'OK'
            ELSE 'Not OK'
        END AS [NAVTOCreationStatus],
        toh.XDock,
        toh.TOPurposeCode,
        ISNULL(dc.LogMsg, '') AS FailedToCreateNAVTO_DC,
        ISNULL(cc.LogMsg, '') AS FailedToCreateNAVTO_CC
    FROM 
        [BCSSoft_ConAppSys].[dbo].[tbTruckLoadHeader] tlh
    LEFT JOIN 
        [BCSSoft_ConAppSys].[dbo].[tbIntgNavLog] dc
    ON 
        tlh.TruckLoadHeaderId = dc.RefHdrId
    AND 
        dc.LogTypeCode = 'ws_CA_TOCounter2DC'
    AND 
        dc.LogStsCode = 'E'
    LEFT JOIN 
        [BCSSoft_ConAppSys].[dbo].[tbIntgNavLog] cc
    ON 
        tlh.TruckLoadHeaderId = cc.RefHdrId
    AND 
        cc.LogTypeCode = 'ws_CA_TOCount2Count'
    AND 
        cc.LogStsCode = 'E'
    LEFT JOIN 
        [BCSSoft_ConAppSys].[dbo].[tbTOHeader] toh
    ON 
        tlh.TONo = toh.TONo
    WHERE 
        tlh.SourceFrom = 'TO'
    AND 
        tlh.TruckLoadStsCode = 'CLOSED'
    AND 
        tlh.CreatedDt >= DATEADD(month, -3, GETDATE())
    AND 
        (dc.LogMsg IS NOT NULL OR cc.LogMsg IS NOT NULL OR tlh.NavTONo NOT LIKE 'EDTO%')
    ORDER BY 
        tlh.CreatedDt DESC;
    """
    df = pd.read_sql(query, conn)
    return df

def generate_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
        workbook = writer.book
        worksheet = writer.sheets['Report']

        # Header formatting
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'align': 'center'
        })
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Autosize columns
        for col_num, col in enumerate(df.columns):
            max_length = max(df[col].astype(str).map(len).max(), len(col))
            worksheet.set_column(col_num, col_num, max_length + 2)

        # Conditional formatting on NAVTOCreationStatus
        navto_col_idx = df.columns.get_loc('NAVTOCreationStatus')

        green_fill = workbook.add_format({'bg_color': '#5CB85C', 'font_color': 'white', 'bold': True})
        red_fill = workbook.add_format({'bg_color': '#D9534F', 'font_color': 'white', 'bold': True})

        for row_num in range(1, len(df) + 1):
            navto_value = df.iloc[row_num - 1]['NAVTOCreationStatus']
            if navto_value == 'OK':
                worksheet.write(row_num, navto_col_idx, navto_value, green_fill)
            elif navto_value == 'Not OK':
                worksheet.write(row_num, navto_col_idx, navto_value, red_fill)

    output.seek(0)
    return output

def main():
    st.title("NAV TO Creation Status Report")

    if st.button("Generate Report"):
        try:
            conn = get_db_connection()
            df = run_query(conn)
            conn.close()

            if df.empty:
                st.warning("No data found matching the criteria.")
                return

            st.success(f"Loaded {len(df)} records.")
            st.dataframe(df)

            excel_data = generate_excel(df)

            st.download_button(
                label="Download Excel Report",
                data=excel_data,
                file_name=f"CA_NAVTO_Creation_Report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")

if __name__ == "__main__":
    main()
