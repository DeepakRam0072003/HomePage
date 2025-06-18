import streamlit as st

st.title("Application Hub")
st.write("Navigate to different modules:")

pages = {
    "C2C_C2D": "pages/C2C_C2D.py",
    "CANAVORP PTSPTR": "pages/CANAVORPPTSPTR.py",
    "CANAVTO PTSPTR": "pages/CANAVTOPTSPTR.py",
    
}

for page_name, page_path in pages.items():
    if st.button(page_name):
        st.switch_page(page_path)
