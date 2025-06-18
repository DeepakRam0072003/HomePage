import streamlit as st

st.title("Application Hub")
st.write("Navigate to different modules:")

pages = {
    "C2C_C2D": "pages/C2C_C2D.py",
    "CANAVORP PTSPTR": "pages/CANAVORPPTSPTR.py",
    "CANAVTO PTSPTR": "pages/CANAVTOPTSPTR.py",
    "ComboCANAVTO CANAVORPstreamlit.py",
    "D2CORPStreamlit.py",
    "SalesErrorLogVSNAVStreamlit.py",
    "SalesReturnErrorLogVSNAVStreamlit.py"
    "StockTakeAdjStreamlit.py",
    "TL TU REStreamlit.py",
    "TO ILE RES Steamlit.py",
    "TO ILE RES Steamlit2.py"
    
}

for page_name, page_path in pages.items():
    if st.button(page_name):
        st.switch_page(page_path)
