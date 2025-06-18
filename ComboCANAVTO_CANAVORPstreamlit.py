import streamlit as st
import pandas as pd
from datetime import datetime
import os
from pathlib import Path
import matplotlib.pyplot as plt
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus
import plotly.express as px
import tempfile

# Configuration for all databases
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
    'orp': {
        'server': 'caappsdb,1435',
        'database': 'BCSSoft_ConAppSys',
        'username': 'Deepak',
        'password': 'Deepak@321',
        'driver': 'ODBC Driver 17 for SQL Server'
    }
}

# Streamlit page configuration
st.set_page_config(
    page_title="Transfer & Truck Load Analysis",
    page_icon="🚛",
    layout="wide"
)

# Custom CSS styling
st.markdown("""
<style>
    .metric-card {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 10px;
    }
    .metric-title {
        font-size: 14px;
        color: #555;
    }
    .metric-value {
        font-size: 24px;
        font-weight: bold;
    }
    .stDataFrame {
        width: 100%;
    }
    .success-text { color: #4CAF50; }
    .warning-text { color: #FF9800; }
    .error-text { color: #F44336; }
</style>
""", unsafe_allow_html=True)

def get_db_connection(db_type):
    """Establish database connection with error handling"""
    try:
        config = CONFIG[db_type]
        connection_string = (
            f"mssql+pyodbc://{config['username']}:%s@{config['server']}/{config['database']}"
            f"?driver={quote_plus(config['driver'])}"
            f"&TrustServerCertificate=yes"
        ) % quote_plus(config['password'])
        engine = create_engine(connection_string)
        # Test connection
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        return engine
    except Exception as e:
        st.error(f"{db_type.upper()} connection failed: {str(e)}")
        return None

def get_transfer_data(engine):
    """Get transfer data from NAV"""
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
        [EDLIVE].[dbo].[Eastern Decorator Sdn_ Bhd_$Transfer Header]
    WHERE 
        Status IN (0, 1)
        AND [Posting Date] >= DATEADD(month, -3, GETDATE())
    """)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(query, conn)
            if df.empty:
                st.warning("No transfer data found")
            return df
    except Exception as e:
        st.error(f"Transfer query failed: {str(e)}")
        return pd.DataFrame()

def get_transfer_shipment_headers(engine):
    """Get shipment headers from NAV"""
    query = text("""
    SELECT 
        [Transfer Order No_],
        [External Document No_],
        [External Document No_ 2]
    FROM 
        [EDLIVE].[dbo].[Eastern Decorator Sdn_ Bhd_$Transfer Shipment Header]
    WHERE
        [Posting Date] >= DATEADD(month, -3, GETDATE())
    """)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(query, conn)
            df['External Document No_'] = df['External Document No_'].str.strip()
            return df
    except Exception as e:
        st.error(f"Transfer shipment query failed: {str(e)}")
        return pd.DataFrame()

def get_transfer_receipt_headers(engine):
    """Get receipt headers from NAV"""
    query = text("""
    SELECT 
        [Transfer Order No_],
        [External Document No_],
        [External Document No_ 2]
    FROM 
        [EDLIVE].[dbo].[Eastern Decorator Sdn_ Bhd_$Transfer Receipt Header]
    WHERE
        [Posting Date] >= DATEADD(month, -3, GETDATE())
    """)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(query, conn)
            df['External Document No_'] = df['External Document No_'].str.strip()
            return df
    except Exception as e:
        st.error(f"Transfer receipt query failed: {str(e)}")
        return pd.DataFrame()

def get_truck_load_errors(engine, source_type):
    """Get truck load errors from TO/ORP"""
    if source_type == 'TO':
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
    else:  # ORP
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
            AND tl.SourceFrom = 'ORP'
            AND (ship.logid IS NOT NULL OR receipt.logid IS NOT NULL)
            AND tl.CreatedDt >= DATEADD(month, -3, GETDATE())
        """)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(query, conn)
            if df.empty:
                st.warning(f"No truck load data found for {source_type}")
            return df
    except Exception as e:
        st.error(f"Truck load query failed for {source_type}: {str(e)}")
        return pd.DataFrame()

def join_and_analyze_data(transfer_df, truck_df, shipment_df, receipt_df, source_type):
    """Combine and analyze datasets with validation"""
    try:
        if source_type == 'TO':
            transfer_df['TONo'] = transfer_df['TONo'].astype(str).str.strip().str.upper()
            truck_df['HostHeaderNo'] = truck_df['HostHeaderNo'].astype(str).str.strip().str.upper()
            
            combined = pd.merge(
                transfer_df,
                truck_df,
                left_on='TONo',
                right_on='HostHeaderNo',
                how='inner',
                suffixes=('_NAV', '_TO')
            )
            
            shipment_df['External Document No_'] = shipment_df['External Document No_'].str.strip().str.upper()
            receipt_df['External Document No_'] = receipt_df['External Document No_'].str.strip().str.upper()
            
            combined['IsPostedShipment'] = combined['HostHeaderNo'].isin(shipment_df['External Document No_'])
            combined['IsPostedReceipt'] = combined['HostHeaderNo'].isin(receipt_df['External Document No_'])
        else:  # ORP
            transfer_df['TONo'] = transfer_df['TONo'].astype(str).str.strip().str.lower()
            truck_df['TONo'] = truck_df['TONo'].astype(str).str.strip().str.lower()
            
            combined = pd.merge(
                transfer_df,
                truck_df,
                on='TONo',
                how='inner',
                suffixes=('_Transfer', '_ORP')
            )
            
            combined['IsPostedReceipt'] = combined['ReceiptErrorMsg'].isna()
            combined['IsPostedShipment'] = combined['ShipErrorMsg'].isna()
        
        combined['IsFullyPosted'] = combined['IsPostedReceipt'] & combined['IsPostedShipment']
        combined['Posting Date'] = pd.to_datetime(combined['Posting Date'])
        combined['LoadClosedDt'] = pd.to_datetime(combined['LoadClosedDt'])
        combined['ProcessingTimeDays'] = (combined['LoadClosedDt'] - combined['Posting Date']).dt.days
        combined['HasError'] = combined['ShipErrorMsg'].notna() | combined['ReceiptErrorMsg'].notna()
        
        return combined.drop_duplicates()
    except Exception as e:
        st.error(f"Data joining failed: {str(e)}")
        return pd.DataFrame()

def display_metrics(df):
    """Display key metrics in cards"""
    st.subheader("Key Metrics")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-title">Total Transfers</div>
            <div class="metric-value">{:,}</div>
        </div>
        """.format(len(df)), unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-title">With Errors</div>
            <div class="metric-value">{:,}</div>
        </div>
        """.format(df['HasError'].sum()), unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-title">Fully Posted</div>
            <div class="metric-value">{:,}</div>
        </div>
        """.format(df['IsFullyPosted'].sum()), unsafe_allow_html=True)
    
    with col4:
        avg_time = df['ProcessingTimeDays'].mean()
        st.markdown("""
        <div class="metric-card">
            <div class="metric-title">Avg Process Time</div>
            <div class="metric-value">{:.1f} days</div>
        </div>
        """.format(avg_time if not pd.isna(avg_time) else 0), unsafe_allow_html=True)

def create_visualizations(df):
    """Create interactive visualizations"""
    st.subheader("Data Visualizations")
    
    tab1, tab2, tab3 = st.tabs(["Status Distribution", "Processing Times", "Error Analysis"])
    
    with tab1:
        status_counts = df['Status Description'].value_counts().reset_index()
        fig = px.pie(status_counts, names='Status Description', values='count', 
                    title='Transfer Status Distribution')
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        fig = px.histogram(df, x='ProcessingTimeDays', 
                         title='Processing Time Distribution (Days)',
                         nbins=20)
        st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        if df['HasError'].sum() > 0:
            error_df = df[df['HasError']]
            error_counts = error_df[['ShipErrorMsg', 'ReceiptErrorMsg']].apply(lambda x: x.notna().sum())
            fig = px.bar(error_counts, title='Error Type Distribution',
                        labels={'index': 'Error Type', 'value': 'Count'})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No errors found in the selected data")

def generate_excel_report(df):
    """Generate Excel report and return as bytes"""
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        with pd.ExcelWriter(tmp.name, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Transfer Data', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Transfer Data']
            
            # Format headers
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top',
                'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
            })
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Auto-adjust column widths
            for idx, col in enumerate(df.columns):
                max_len = max((df[col].astype(str).map(len).max(), len(str(col)))) + 2
                worksheet.set_column(idx, idx, min(max_len, 30))
        
        with open(tmp.name, 'rb') as f:
            excel_data = f.read()
    
    os.unlink(tmp.name)
    return excel_data

def run_analysis(source_type):
    """Run the complete analysis workflow"""
    with st.spinner("Connecting to databases..."):
        nav_engine = get_db_connection('nav')
        if source_type == 'TO':
            target_engine = get_db_connection('to')
        else:
            target_engine = get_db_connection('orp')
        
        if not nav_engine or not target_engine:
            return None
    
    progress_bar = st.progress(0)
    
    with st.spinner("Fetching transfer data..."):
        transfer_df = get_transfer_data(nav_engine)
        progress_bar.progress(20)
    
    with st.spinner("Fetching shipment data..."):
        shipment_df = get_transfer_shipment_headers(nav_engine)
        progress_bar.progress(40)
    
    with st.spinner("Fetching receipt data..."):
        receipt_df = get_transfer_receipt_headers(nav_engine)
        progress_bar.progress(60)
    
    with st.spinner(f"Fetching {source_type} truck load data..."):
        truck_df = get_truck_load_errors(target_engine, source_type)
        progress_bar.progress(80)
    
    with st.spinner("Analyzing data..."):
        combined_df = join_and_analyze_data(transfer_df, truck_df, shipment_df, receipt_df, source_type)
        progress_bar.progress(100)
    
    # Clean up connections
    nav_engine.dispose()
    target_engine.dispose()
    
    return combined_df

def main():
    st.title("🚛 Transfer & Truck Load Analysis")
    
    # Initialize session state
    if 'report_data' not in st.session_state:
        st.session_state.report_data = None
    if 'source_type' not in st.session_state:
        st.session_state.source_type = None
    
    # Sidebar controls
    with st.sidebar:
        st.header("Analysis Settings")
        source_type = st.radio("Select Data Source", ['TO', 'ORP'], index=0)
        
        if st.button("Run Analysis"):
            st.session_state.source_type = source_type
            st.session_state.report_data = run_analysis(source_type)
    
    # Display results
    if st.session_state.report_data is not None:
        df = st.session_state.report_data
        
        if not df.empty:
            st.success(f"Analysis completed for {st.session_state.source_type} data")
            display_metrics(df)
            create_visualizations(df)
            
            st.subheader("Detailed Data")
            st.dataframe(df)
            
            # Export options
            st.subheader("Export Data")
            col1, col2 = st.columns(2)
            
            with col1:
                excel_data = generate_excel_report(df)
                st.download_button(
                    label="Download Excel Report",
                    data=excel_data,
                    file_name=f"{st.session_state.source_type}_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name=f"{st.session_state.source_type}_data.csv",
                    mime="text/csv"
                )
        else:
            st.warning("No data available for the selected criteria")

if __name__ == "__main__":
    main()