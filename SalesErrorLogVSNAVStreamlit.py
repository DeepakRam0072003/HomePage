import streamlit as st
import pandas as pd
import pyodbc
from datetime import datetime
from dateutil.relativedelta import relativedelta
from io import BytesIO

def create_conn_str(config):
    return (
        f"DRIVER={{{config['driver']}}};"
        f"SERVER={config['server']};"
        f"DATABASE={config['database']};"
        f"UID={config['username']};"
        f"PWD={config['password']};"
        f"Timeout={config.get('timeout', 30)};"
    )

def fetch_bcs_data(start_date, end_date, conn_str):
    query = """
    WITH RankedRows AS (
        SELECT 
            h.SOHeaderId,
            h.SONo,
            h.SOTypeCode,
            CONVERT(varchar, h.SODt, 120) AS SODt,
            d.SODetailId,
            l.LogTypeCode,
            l.LogStsCode,
            l.LogMsg,
            ROW_NUMBER() OVER (PARTITION BY h.SOHeaderId ORDER BY d.SODetailId) AS rn
        FROM 
            [BCSSoft_ConAppSys].[dbo].[tbSOHeader] h
        INNER JOIN 
            [BCSSoft_ConAppSys].[dbo].[tbSODetail] d ON h.SOHeaderId = d.SOHeaderId
        INNER JOIN 
            [BCSSoft_ConAppSys].[dbo].[tbIntgNavLog] l ON h.SOHeaderId = l.RefHdrId AND d.SODetailId = l.RefDtlId
        WHERE 
            l.LogTypeCode = 'ws_ItemJournal'
            AND l.LogStsCode = 'E'
            AND h.SODt BETWEEN ? AND ?
    )
    SELECT 
        SOHeaderId,
        SONo,
        SOTypeCode,
        SODt,
        SODetailId,
        LogTypeCode,
        LogStsCode,
        LogMsg
    FROM RankedRows
    WHERE rn = 1
    ORDER BY SOHeaderId
    """
    with pyodbc.connect(conn_str) as conn:
        df = pd.read_sql(query, conn, params=[start_date, end_date])
    return df

def sono_to_docno(sono):
    try:
        if not isinstance(sono, str) or not sono.startswith("SO_"):
            return None
        parts = sono[3:].split("_")
        if len(parts) != 3:
            return None
        prefix, date, time = parts
        short_date = date[2:] if date.startswith("20") else date
        return f"{prefix}{short_date}{time}"
    except Exception:
        return None

def fetch_existing_docnos(doc_no_list, conn_str):
    if not doc_no_list:
        return set()
    in_clause = ",".join(f"'{doc}'" for doc in doc_no_list)
    query = f"""
    SELECT DISTINCT [Document No_]
    FROM [EDLIVE].[dbo].[Eastern Decorator Sdn_ Bhd_$Item Ledger Entry]
    WHERE [Entry Type] = '1'
    AND [Document No_] IN ({in_clause})
    """
    with pyodbc.connect(conn_str) as conn:
        df = pd.read_sql(query, conn)
    return set(df['Document No_'].dropna().unique())

def create_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
        workbook = writer.book
        worksheet = writer.sheets['Report']

        # Define formats
        header_format = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white'})
        green_format = workbook.add_format({'font_color': 'green'})
        red_format = workbook.add_format({'font_color': 'red'})

        # Apply header format and auto-adjust column width
        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, header_format)
            max_width = max(df[value].astype(str).map(len).max(), len(value)) + 2
            worksheet.set_column(col_num, col_num, max_width)

        # Apply color formatting to IsPostedILE column
        if 'IsPostedILE' in df.columns:
            col_index = df.columns.get_loc('IsPostedILE')
            for row_num, value in enumerate(df['IsPostedILE'], start=1):
                fmt = green_format if value == 'OK' else red_format
                worksheet.write(row_num, col_index, value, fmt)
    output.seek(0)
    return output

def main():
    st.title("BCS SO Log Error Report")

    # Connection configs
    bcs = {
        'server': 'caappsdb,1435',
        'database': 'BCSSoft_ConAppSys',
        'username': 'Deepak',
        'password': 'Deepak@321',
        'driver': 'ODBC Driver 17 for SQL Server',
        'timeout': 30
    }

    nav = {
        'server': 'nav18db',
        'database': 'EDLIVE',
        'username': 'barcode1',
        'password': 'barcode@1433',
        'driver': 'ODBC Driver 17 for SQL Server',
        'timeout': 30
    }

    # Date range picker (default last 3 months)
    today = datetime.today()
    default_start = today - relativedelta(months=3)

    start_date = st.date_input("Start Date", default_start)
    end_date = st.date_input("End Date", today)

    if start_date > end_date:
        st.error("Start Date must be before or equal to End Date.")
        return

    if st.button("Fetch Data"):
        bcs_conn_str = create_conn_str(bcs)
        nav_conn_str = create_conn_str(nav)

        with st.spinner("Fetching BCS data..."):
            df_bcs = fetch_bcs_data(start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'), bcs_conn_str)

        if df_bcs.empty:
            st.warning("No BCS data found for the selected date range.")
            return

        # Add Document No_ column
        df_bcs['Document No_'] = df_bcs['SONo'].apply(sono_to_docno)

        doc_no_list = df_bcs['Document No_'].dropna().unique().tolist()

        with st.spinner(f"Checking {len(doc_no_list)} Document Nos in NAV..."):
            existing_docnos = fetch_existing_docnos(doc_no_list, nav_conn_str)

        df_bcs['IsPostedILE'] = df_bcs['Document No_'].apply(lambda doc: 'OK' if doc in existing_docnos else 'NotOK')

        # Reorder columns to put 'Document No_' and 'IsPostedILE' before 'LogMsg'
        cols = df_bcs.columns.tolist()
        if 'LogMsg' in cols and 'Document No_' in cols and 'IsPostedILE' in cols:
            logmsg_index = cols.index('LogMsg')
            for col in ['IsPostedILE', 'Document No_']:
                cols.remove(col)
                cols.insert(logmsg_index, col)
            df_bcs = df_bcs[cols]

        st.success(f"Loaded {len(df_bcs)} records.")
        st.dataframe(df_bcs)

        excel_data = create_excel(df_bcs)

        st.download_button(
            label="Download Excel Report",
            data=excel_data,
            file_name=f"SO_Log_Error_Report_BCS_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
