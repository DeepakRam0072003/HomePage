import streamlit as st
import pandas as pd
from datetime import datetime
import os
from pathlib import Path
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus

CONFIG = {
    'nav': {
        'server': 'nav18db',
        'database': 'EDLIVE',
        'username': 'barcode1',
        'password': 'barcode@1433',
        'driver': 'ODBC Driver 17 for SQL Server'
    },
    'orp': {
        'server': 'caappsdb,1435',
        'database': 'BCSSoft_ConAppSys',
        'username': 'Deepak',
        'password': 'Deepak@321',
        'driver': 'ODBC Driver 17 for SQL Server'
    },
    'output_folder': r'\\hq-file01\fileshare\CavsNavErrors\NAV-TransactionErrorLog'
}

def setup_output_folder():
    Path(CONFIG['output_folder']).mkdir(parents=True, exist_ok=True)
    return CONFIG['output_folder']

def get_db_connection(db_type):
    config = CONFIG[db_type]
    connection_string = (
        f"mssql+pyodbc://{config['username']}:%s@{config['server']}/{config['database']}"
        f"?driver={quote_plus(config['driver'])}&TrustServerCertificate=yes"
    ) % quote_plus(config['password'])
    return create_engine(connection_string)

def get_transfer_data(engine):
    query = text("""
    SELECT
        [No_],
        CASE
            WHEN Status = 1 THEN 'Released'
            WHEN Status = 0 THEN 'Open'
            ELSE CAST(Status AS VARCHAR(10))
        END AS [Status Description],
        [External Document No_] AS [TONo],
        [External Document No_ 2],
        [Transfer-from Code],
        [Transfer-to Code],
        [Posting Date]
    FROM
        [EDLIVE].[dbo].[Eastern Decorator Sdn_ Bhd_$Transfer Header]
    WHERE
        Status IN (0, 1)
    """)
    with engine.connect() as conn:
        return pd.read_sql(query, conn)

def get_truck_load_errors(engine):
    query = text("""
    SELECT
        tl.TruckLoadHeaderId,
        tl.TONo,
        tl.HostHeaderNo,
        tl.TruckLoadStsCode,
        tl.LoadClosedDt,
        tl.UnloadClosedDt,
        tl.SourceFrom,
        tl.CreatedDt,
        ship.LogMsg AS ShipErrorMsg,
        receipt.LogMsg AS ReceiptErrorMsg
    FROM
        [BCSSoft_ConAppSys].[dbo].[tbTruckLoadHeader] tl
    LEFT JOIN
        [BCSSoft_ConAppSys].[dbo].[tbIntgNavLog] ship
        ON tl.TONo = ship.WMSOrderKey
        AND ship.LogTypeCode = 'ws_CA_PostTO-TO(PostReceipt)'
        AND ship.LogStsCode = 'E'
    LEFT JOIN
        [BCSSoft_ConAppSys].[dbo].[tbIntgNavLog] receipt
        ON tl.TONo = receipt.WMSOrderKey
        AND receipt.LogTypeCode = 'ws_CA_PostTO-ORP(PostReceipt)'
        AND receipt.LogStsCode = 'E'
    WHERE
        tl.UnloadClosedDt IS NOT NULL
        AND tl.SourceFrom = 'orp'
        AND tl.SourceFrom <> ''
        AND (ship.logid IS NOT NULL OR receipt.logid IS NOT NULL)
    """)
    with engine.connect() as conn:
        return pd.read_sql(query, conn)

def get_shipment_receipt_data(engine):
    shipment_query = text("""
        SELECT 
            [Transfer Order No_] AS TransferOrderNo,
            [External Document No_] AS TONo_Shipment
        FROM [EDLIVE].[dbo].[Eastern Decorator Sdn_ Bhd_$Transfer Shipment Header]
    """)
    receipt_query = text("""
        SELECT 
            [Transfer Order No_] AS TransferOrderNo,
            [External Document No_] AS TONo_Receipt
        FROM [EDLIVE].[dbo].[Eastern Decorator Sdn_ Bhd_$Transfer Receipt Header]
    """)
    with engine.connect() as conn:
        shipment_df = pd.read_sql(shipment_query, conn)
        receipt_df = pd.read_sql(receipt_query, conn)

    shipment_df['TONo_Shipment'] = shipment_df['TONo_Shipment'].astype(str).str.strip().str.lower()
    receipt_df['TONo_Receipt'] = receipt_df['TONo_Receipt'].astype(str).str.strip().str.lower()

    return shipment_df, receipt_df

def join_and_analyze_data(transfer_df, orp_df, shipment_df, receipt_df):
    transfer_df['TONo'] = transfer_df['TONo'].astype(str).str.strip().str.lower()
    orp_df['TONo'] = orp_df['TONo'].astype(str).str.strip().str.lower()

    combined = pd.merge(
        transfer_df,
        orp_df,
        on='TONo',
        how='inner',
        suffixes=('_Transfer', '_ORP')
    )

    combined = combined[combined['SourceFrom'].str.lower() == 'orp']

    combined['IsPostedReceipt'] = combined['ReceiptErrorMsg'].isna()
    combined['IsPostedShipment'] = combined['ShipErrorMsg'].isna()
    combined['IsFullyPosted'] = combined['IsPostedReceipt'] & combined['IsPostedShipment']

    combined.drop(columns=['LoadClosedDt', 'Posting Date'], inplace=True, errors='ignore')

    return combined

def save_report(df, output_folder):
    try:
        report_path = os.path.join(output_folder, 'ORP_Report_Latest.xlsx')
        writer = pd.ExcelWriter(report_path, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='ORP Data', index=False)

        workbook = writer.book
        worksheet = writer.sheets['ORP Data']

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
            worksheet.set_column(col_num, col_num, 20)

        true_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        false_format = workbook.add_format({'bg_color': '#F8CBAD', 'font_color': '#9C0006'})

        for col in ['IsPostedReceipt', 'IsPostedShipment', 'IsFullyPosted']:
            if col in df.columns:
                col_index = df.columns.get_loc(col)
                worksheet.conditional_format(1, col_index, len(df), col_index,
                                             {'type': 'cell', 'criteria': '==', 'value': 'TRUE', 'format': true_format})
                worksheet.conditional_format(1, col_index, len(df), col_index,
                                             {'type': 'cell', 'criteria': '==', 'value': 'FALSE', 'format': false_format})

        summary = workbook.add_worksheet('Summary')
        summary.write('A1', 'Metric', header_format)
        summary.write('B1', 'Value', header_format)

        summary.write('A2', 'Total Records')
        summary.write('B2', len(df))
        summary.write('A3', 'Posted Shipment')
        summary.write('B3', df['IsPostedShipment'].sum())
        summary.write('A4', 'Posted Receipt')
        summary.write('B4', df['IsPostedReceipt'].sum())
        summary.write('A5', 'Fully Posted')
        summary.write('B5', df['IsFullyPosted'].sum())

        writer.close()
        return report_path
    except Exception as e:
        return None

def display_metrics(df):
    st.subheader("📊 Summary Metrics")
    total = len(df)
    posted_shipment = df['IsPostedShipment'].sum()
    posted_receipt = df['IsPostedReceipt'].sum()
    fully_posted = df['IsFullyPosted'].sum()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Records", total)
    col2.metric("Posted Shipment", posted_shipment)
    col3.metric("Posted Receipt", posted_receipt)
    col4.metric("Fully Posted", fully_posted)

def streamlit_app():
    st.title("🚚 ORP Error Report Dashboard")

    with st.spinner("🔄 Loading data..."):
        output_folder = setup_output_folder()
        nav_engine = get_db_connection('nav')
        orp_engine = get_db_connection('orp')

        transfer_df = get_transfer_data(nav_engine)
        orp_df = get_truck_load_errors(orp_engine)
        shipment_df, receipt_df = get_shipment_receipt_data(nav_engine)

        combined_df = join_and_analyze_data(transfer_df, orp_df, shipment_df, receipt_df)

    display_metrics(combined_df)

    st.subheader("📋 Detailed ORP Data")
    st.dataframe(combined_df)

    if st.button("📥 Save Report"):
        report_path = save_report(combined_df, output_folder)
        if report_path:
            st.success(f"Report saved successfully: {report_path}")
        else:
            st.error("Failed to save report.")

if __name__ == "__main__":
    streamlit_app()
