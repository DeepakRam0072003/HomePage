import streamlit as st
import os

st.title("Application Hub")
st.write("Navigate to different modules:")

pages = {
    "C2C_C2D": "Pages/C2C_C2DStreamlit.py",
    "CANAVORPPTSPTR": "Pages/CANAVORPPTSPTRStreamlit.py",
    "CANAVTOPTSPTR": "Pages/CANAVTOPTSPTRStreamlit.py",
    "ComboCANAVTO_CANAVORP": "Pages/ComboCANAVTO_CANAVORPstreamlit.py",
    "D2CORP": "Pages/D2CORPStreamlit.py",
    "SalesErrorLogVSNAV": "Pages/SalesErrorLogVSNAVStreamlit.py",
    "SalesReturnErrorLogVSNAV": "Pages/SalesReturnErrorLogVSNAVStreamlit.py",
    "StockTakeAdj": "Pages/StockTakeAdjSteamlit.py",  # Note: Steamlit vs Streamlit
    "TL_TU_RE": "Pages/TL_TU_REStreamlit.py",
    "TO_ILE_RES": "Pages/TO_ILE_RES_Streamlit.py",
    "TO_ILE_RES_V2": "Pages/TO_ILE_RES_Streamlit2.py"
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
