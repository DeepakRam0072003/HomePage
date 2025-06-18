import streamlit as st
import os

st.title("Application Hub")
st.write("Navigate to different modules:")

# Debug current directory
st.write("Current directory:", os.getcwd())
st.write("Directory contents:", os.listdir())

# Dictionary with direct file paths (no 'pages/' prefix)
pages = {
    "C2C_C2D": "C2C_C2DStreamlit.py",
    "CANAVORPPTSPTR": "CANAVORPPTSPTRStreamlit.py",
    "CANAVTOPTSPTR": "CANAVTOPTSPTRStreamlit.py",
    "ComboCANAVTO_CANAVORP": "ComboCANAVTO_CANAVORPstreamlit.py",
    "D2CORP": "D2CORPStreamlit.py",
    "SalesErrorLogVSNAV": "SalesErrorLogVSNAVStreamlit.py",
    "SalesReturnErrorLogVSNAV": "SalesReturnErrorLogVSNAVStreamlit.py",
    "StockTakeAdj": "StockTakeAdjStreamlit.py",
    "TL_TU_RE": "TL_TU_REStreamlit.py",
    "TO_ILE_RES": "TO_ILE_RES_Steamlit.py",
    "TO_ILE_RES_V2": "TO_ILE_RES_Steamlit2.py"
}

for page_name, page_file in pages.items():
    if st.button(page_name):
        if os.path.exists(page_file):
            st.switch_page(page_file)
        else:
            st.error(f"File not found: {page_file}")
