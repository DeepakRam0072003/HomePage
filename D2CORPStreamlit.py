import streamlit as st
import pandas as pd
import pyodbc
from io import BytesIO

# DB connection parameters
server = 'caappsdb,1435'
database = 'BCSSoft_ConAppSys'
username = 'Deepak'
password = 'Deepak@321'

def get_connection():
    conn_str = (
        f"DRIVER={{SQL Server}};"
        f"SERVER={server};DATABASE={database};UID={username};PWD={password}"
    )
    conn = pyodbc.connect(conn_str)
    return conn

def fetch_orp_data(conn):
    query = """
    SELECT DISTINCT
        orph.CreatedDt, 
        orph.ORPTempHdrId,
        orph.ORPNo,
        orph.CounterCode AS ShipToCounter,
        orph.ORPStatus,
        orpd.NavTONo,
        CASE 
            WHEN orpd.NavTONo IS NOT NULL AND orpd.NavTONo <> '' THEN 'OK'
            ELSE 'Not OK'
        END AS NAVTOCreationStatus,
        orpd.wmsorderkey,
        orpd.WMSCfmSts,
        ln.LogMsg AS FailedToCreateNAVTO
    FROM 
        [BCSSoft_ConAppSys].[dbo].[tbORPTempHdr] orph
    JOIN 
        [BCSSoft_ConAppSys].[dbo].[tbORPTempDtl] orpd
    ON 
        orph.ORPTempHdrId = orpd.ORPTempHdrId
    LEFT JOIN 
        [BCSSoft_ConAppSys].[dbo].[tbIntgNavLog] ln
    ON 
        orpd.wmsorderkey = ln.wmsorderkey
    AND 
        ln.LogTypeCode = 'ws_CA_TODc2Counter'
    AND 
        ln.LogStsCode = 'E'
    WHERE
        orph.ORPStatus = 'WMSShipped'
    AND
        orpd.WMSCfmSts = 'Shipped Complete'
    AND
        ln.LogMsg IS NOT NULL
    AND
        orph.CreatedDt >= DATEADD(month, -3, GETDATE())  -- Last 3 months
    ORDER BY 
        orph.CreatedDt DESC;
    """
    df = pd.read_sql(query, conn)
    df.drop_duplicates(subset=['wmsorderkey'], inplace=True)
    return df

def create_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
        workbook = writer.book
        worksheet = writer.sheets['Report']

        # Header format
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

        # Conditional formatting for NAVTOCreationStatus
        red_fill = workbook.add_format({'bg_color': '#FF0000', 'font_color': 'white'})
        green_fill = workbook.add_format({'bg_color': '#00FF00', 'font_color': 'black'})

        navto_col_idx = df.columns.get_loc('NAVTOCreationStatus')

        for row_num in range(1, len(df) + 1):
            navto_value = df.iloc[row_num - 1]['NAVTOCreationStatus']
            if navto_value == 'Not OK':
                worksheet.write(row_num, navto_col_idx, navto_value, red_fill)
            elif navto_value == 'OK':
                worksheet.write(row_num, navto_col_idx, navto_value, green_fill)

        # Auto-adjust column widths
        for col_num, column in enumerate(df.columns):
            max_length = max(df[column].astype(str).apply(len).max(), len(str(column)))
            worksheet.set_column(col_num, col_num, max_length + 2)

    output.seek(0)
    return output

def main():
    st.title("ORP NAVTO Creation Report")

    if st.button("Generate Report"):
        try:
            conn = get_connection()
            df = fetch_orp_data(conn)
            conn.close()

            if df.empty:
                st.warning("No data found for the last 3 months.")
                return

            st.success(f"Loaded {len(df)} records.")
            st.dataframe(df)

            excel_bytes = create_excel(df)

            st.download_button(
                label="Download Excel Report",
                data=excel_bytes,
                file_name=f"CA_NAVORP_Creation_Report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")

if __name__ == "__main__":
    main()
