import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus
from io import BytesIO

# Configuration for both databases
CONFIG = {
    'nav': {
        'server': 'nav18db',
        'database': 'EDLIVE',
        'username': 'barcode1',
        'password': 'barcode@1433',
        'driver': 'ODBC Driver 17 for SQL Server'
    },
    'to': {
        'server': 'caappsdb,1435',
        'database': 'BCSSoft_ConAppSys',
        'username': 'Deepak',
        'password': 'Deepak@321',
        'driver': 'ODBC Driver 17 for SQL Server'
    },
    'output_folder': r'\\hq-file01\fileshare\CavsNavErrors\NAV-TransactionErrorLog'  # Not used for Streamlit download
}

def get_db_connection(db_type):
    try:
        config = CONFIG[db_type]
        connection_string = (
            f"mssql+pyodbc://{config['username']}:%s@{config['server']}/{config['database']}"
            f"?driver={quote_plus(config['driver'])}"
            f"&TrustServerCertificate=yes"
        ) % quote_plus(config['password'])
        engine = create_engine(connection_string)
        return engine
    except Exception as e:
        st.error(f"🚨 {db_type.upper()} connection failed: {str(e)}")
        raise

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
        [External Document No_ 2] AS [TONo2],
        [Transfer-from Code],
        [Transfer-to Code],
        [Posting Date]
    FROM 
        [dbo].[Eastern Decorator Sdn_ Bhd_$Transfer Header]
    WHERE 
        Status IN (0, 1)
        AND [Posting Date] >= DATEADD(month, -3, GETDATE())
    """)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(query, conn)
            if 'TONo' not in df.columns:
                raise ValueError("TONo column not found in transfer data")
            return df
    except Exception as e:
        st.error(f"🚨 Transfer query failed: {str(e)}")
        raise

def get_transfer_shipment_headers(engine):
    query = text("""
    SELECT 
        [Transfer Order No_],
        [External Document No_],
        [External Document No_ 2]
    FROM 
        [dbo].[Eastern Decorator Sdn_ Bhd_$Transfer Shipment Header]
    WHERE
        [Posting Date] >= DATEADD(month, -3, GETDATE())
    """)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(query, conn)
            df['External Document No_'] = df['External Document No_'].str.strip()
            return df
    except Exception as e:
        st.error(f"🚨 Transfer shipment query failed: {str(e)}")
        raise

def get_transfer_receipt_headers(engine):
    query = text("""
    SELECT 
        [Transfer Order No_],
        [External Document No_],
        [External Document No_ 2]
    FROM 
        [dbo].[Eastern Decorator Sdn_ Bhd_$Transfer Receipt Header]
    WHERE
        [Posting Date] >= DATEADD(month, -3, GETDATE())
    """)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(query, conn)
            df['External Document No_'] = df['External Document No_'].str.strip()
            return df
    except Exception as e:
        st.error(f"🚨 Transfer receipt query failed: {str(e)}")
        raise

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
        ON (tl.TONo = ship.WMSOrderKey OR tl.TruckLoadHeaderId = ship.RefHdrId)
        AND ship.LogTypeCode = 'ws_CA_PostTO-TO(PostShip)' 
        AND ship.LogStsCode = 'E'
    LEFT JOIN 
        [BCSSoft_ConAppSys].[dbo].[tbIntgNavLog] receipt
        ON (tl.TONo = receipt.WMSOrderKey OR tl.TruckLoadHeaderId = receipt.RefHdrId)
        AND receipt.LogTypeCode = 'ws_CA_PostTO-TO(PostReceipt)' 
        AND receipt.LogStsCode = 'E'
    WHERE 
        tl.UnloadClosedDt IS NOT NULL 
        AND tl.SourceFrom = 'TO'
        AND (ship.logid IS NOT NULL OR receipt.logid IS NOT NULL)
        AND tl.CreatedDt >= DATEADD(month, -3, GETDATE())
    """)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(query, conn)
            if 'HostHeaderNo' not in df.columns:
                raise ValueError("HostHeaderNo column not found in TO data")
            return df
    except Exception as e:
        st.error(f"🚨 Truck load query failed: {str(e)}")
        raise

def join_and_analyze_data(transfer_df, to_df, shipment_df, receipt_df):
    try:
        # Check required columns
        if 'TONo' not in transfer_df.columns:
            raise ValueError("Missing required column: TONo in transfer_df")
        if 'HostHeaderNo' not in to_df.columns:
            raise ValueError("Missing required column: HostHeaderNo in to_df")

        # Normalize keys for join
        transfer_df['TONo'] = transfer_df['TONo'].astype(str).str.strip().str.upper()
        to_df['HostHeaderNo'] = to_df['HostHeaderNo'].astype(str).str.strip().str.upper()

        # Merge NAV transfer with TO truck load
        combined = pd.merge(
            transfer_df,
            to_df,
            left_on='TONo',
            right_on='HostHeaderNo',
            how='inner',
            suffixes=('_NAV', '_TO')
        )

        if combined.empty:
            st.warning("⚠️ Warning: No records matched after merge")
            return pd.DataFrame()

        # Prepare shipment and receipt external document keys
        shipment_df['External Document No_'] = shipment_df['External Document No_'].str.strip().str.upper()
        receipt_df['External Document No_'] = receipt_df['External Document No_'].str.strip().str.upper()

        # Add posting flags
        combined['IsPostedShipment'] = combined['HostHeaderNo'].isin(shipment_df['External Document No_'])
        combined['IsPostedReceipt'] = combined['HostHeaderNo'].isin(receipt_df['External Document No_'])
        combined['IsFullyPosted'] = combined['IsPostedShipment'] & combined['IsPostedReceipt']

        return combined.drop_duplicates()
    except Exception as e:
        st.error(f"🚨 Data joining failed: {str(e)}")
        raise

def save_report_to_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Transfer Data', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Transfer Data']

        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
        })
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for idx, col in enumerate(df.columns):
            max_len = max((df[col].astype(str).map(len).max(), len(str(col)))) + 2
            worksheet.set_column(idx, idx, min(max_len, 30))

        # Conditional formatting for boolean columns
        bool_cols = ['IsPostedShipment', 'IsPostedReceipt', 'IsFullyPosted']
        for col_name in bool_cols:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name)
                worksheet.conditional_format(1, col_idx, len(df), col_idx, {
                    'type': 'cell',
                    'criteria': '==',
                    'value': True,
                    'format': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
                })
                worksheet.conditional_format(1, col_idx, len(df), col_idx, {
                    'type': 'cell',
                    'criteria': '==',
                    'value': False,
                    'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                })

        summary_sheet = workbook.add_worksheet('Summary')
        summary_sheet.set_column('A:A', 30)
        summary_sheet.set_column('B:B', 20)
        metrics = [
            ['Total Transfers', len(df)],
            ['Posted Shipments', df['IsPostedShipment'].sum() if 'IsPostedShipment' in df.columns else 0],
            ['Posted Receipts', df['IsPostedReceipt'].sum() if 'IsPostedReceipt' in df.columns else 0],
            ['Fully Posted', df['IsFullyPosted'].sum() if 'IsFullyPosted' in df.columns else 0],
            ['Report Generated', datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
        ]
        bold_format = workbook.add_format({'bold': True})
        for row_num, (label, value) in enumerate(metrics):
            summary_sheet.write(row_num, 0, label, bold_format)
            summary_sheet.write(row_num, 1, str(value))

    output.seek(0)
    return output

def main():
    st.title("📊 NAV-TO Transfer and Truck Load Error Report")
    st.write("Connects to NAV and TO databases, fetches recent transfer data and truck load errors, then displays and exports combined analysis.")

    if st.button("Generate Report"):
        try:
            with st.spinner("Connecting to NAV database..."):
                nav_engine = get_db_connection('nav')
            with st.spinner("Fetching Transfer Header data..."):
                transfer_df = get_transfer_data(nav_engine)
            with st.spinner("Fetching Transfer Shipment Header data..."):
                shipment_df = get_transfer_shipment_headers(nav_engine)
            with st.spinner("Fetching Transfer Receipt Header data..."):
                receipt_df = get_transfer_receipt_headers(nav_engine)
            with st.spinner("Connecting to TO database..."):
                to_engine = get_db_connection('to')
            with st.spinner("Fetching Truck Load error data from TO..."):
                to_df = get_truck_load_errors(to_engine)

            st.success("Data loaded successfully! Merging datasets...")

            combined_df = join_and_analyze_data(transfer_df, to_df, shipment_df, receipt_df)

            if combined_df.empty:
                st.warning("No data to display after merging. Try changing filters or check DB status.")
                return

            st.dataframe(combined_df)

            excel_bytes = save_report_to_bytes(combined_df)

            st.download_button(
                label="📥 Download Excel Report",
                data=excel_bytes,
                file_name=f"TO_Transfer_Error_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error occurred: {e}")

        finally:
            if 'nav_engine' in locals():
                nav_engine.dispose()
            if 'to_engine' in locals():
                to_engine.dispose()

if __name__ == "__main__":
    main()
