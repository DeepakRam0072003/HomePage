import streamlit as st
import os  # Added for debugging

st.title("Application Hub")
st.write("Navigate to different modules:")

pages = {
    "C2C_C2D": "pages/C2C_C2DStreamlit.py",
    "CANAVORPPTSPTR": "pages/CANAVORPPTSPTRStreamlit.py",
    "CANAVTOPTSPTR": "pages/CANAVTOPTSPTRStreamlit.py",
    "ComboCANAVTO_CANAVORP": "pages/ComboCANAVTO_CANAVORPstreamlit.py",  # Fixed key
    "D2CORP": "pages/D2CORPStreamlit.py",
    "SalesErrorLogVSNAV": "pages/SalesErrorLogVSNAVStreamlit.py",
    "SalesReturnErrorLogVSNAV": "pages/SalesReturnErrorLogVSNAVStreamlit.py",
    "StockTakeAdj": "pages/StockTakeAdjStreamlit.py",
    "TL_TU_RE": "pages/TL_TU_REStreamlit.py",
    "TO_ILE_RES": "pages/TO_ILE_RES_Streamlit.py",
    "TO_ILE_RES_V2": "pages/TO_ILE_RES_Streamlit2.py"  # Changed key and fixed typo
}

# Debugging: Show available files
st.write("Files in pages directory:", os.listdir("pages"))

for page_name, page_path in pages.items():
    if st.button(page_name):
        if os.path.exists(page_path):
            st.switch_page(page_path)
        else:
            st.error(f"File not found: {page_path}")
