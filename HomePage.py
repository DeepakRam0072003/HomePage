import streamlit as st
import os

st.title("Application Hub")
st.write("Navigate to different modules:")

pages = {
    "C2C_C2D": "pages/C2C_C2DStreamlit.py",
    "CANAVORPPTSPTR": "pages/CANAVORPPTSPTRStreamlit.py",
    "CANAVTOPTSPTR": "pages/CANAVTOPTSPTRStreamlit.py",
    "ComboCANAVTO_CANAVORP": "pages/ComboCANAVTO_CANAVORPstreamlit.py",
    "D2CORP": "pages/D2CORPStreamlit.py",
    "SalesErrorLogVSNAV": "pages/SalesErrorLogVSNAVStreamlit.py",
    "SalesReturnErrorLogVSNAV": "pages/SalesReturnErrorLogVSNAVStreamlit.py",
    "StockTakeAdj": "pages/StockTakeAdjStreamlit.py",  # Note: Steamlit vs Streamlit
    "TL_TU_RE": "pages/TL_TU_REStreamlit.py",
    "TO_ILE_RES": "pages/TO_ILE_RES_Streamlit.py",
    "TO_ILE_RES_V2": "pages/TO_ILE_RES_Streamlit2.py"
}

# Debug current structure
try:
    st.write("Current directory:", os.getcwd())
    st.write("Directory contents:", os.listdir())
    st.write("Pages contents:", os.listdir("Pages"))  # Capital 'P' to match your structure
except FileNotFoundError as e:
    st.error(f"Directory error: {str(e)}")

# Navigation buttons
for page_name, page_path in pages.items():
    if st.button(page_name):
        if os.path.exists(page_path):
            st.switch_page(page_path)
        else:
            st.error(f"File not found: {page_path}")
