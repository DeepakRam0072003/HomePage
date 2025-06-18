import streamlit as st

st.title("Application Hub")
st.write("Navigate to different modules:")

pages = {
    "C2C_C2D": "pages/C2C_C2D.py",
    "CANAVORPPTSPTR": "pages/CANAVORPPTSPTR.py",
    "CANAVTOPTSPTR": "pages/CANAVTOPTSPTR.py",
    "ComboCANAVTO_CANAVORPS": "pages/ComboCANAVTO_CANAVORPStreamlit.py",
    "D2CORP": "pages/D2CORPStreamlit.py",
    "SalesErrorLogVSNAV": "pages/SalesErrorLogVSNAVStreamlit.py",
    "SalesReturnErrorLogVSNAV": "pages/SalesReturnErrorLogVSNAVStreamlit.py",
    "StockTakeAdj": "pages/StockTakeAdjStreamlit.py",
    "TL_TU_RE": "pages/TL_TU_REStreamlit.py",
    "TO_ILE_RES": "pages/TO_ILE_RES_Steamlit.py",
    "TO_ILE_RES": "pages/TO_ILE_RES_Steamlit2.py"
    
}

for page_name, page_path in pages.items():
    if st.button(page_name):
        st.switch_page(page_path)
