import streamlit as st

st.set_page_config(
    page_title="WIN-SQUARE | Requirement Sheet Engine",
    layout="wide",
    page_icon="🪟",
    initial_sidebar_state="expanded"
)
main_page = st.Page("app_home.py", title="Requirement Sheet Engine", icon="🪟", default=True)
frame_page = st.Page("pages/1_Frame_Details.py", title="Frame Size (WxH) & SQFT", icon="🖼️")

pg = st.navigation({"Main Menu": [main_page], "Tools": [frame_page]})
pg.run()
