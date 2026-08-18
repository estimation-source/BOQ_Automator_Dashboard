import streamlit as st


main_page = st.Page("app_home.py", title="Requirement Sheet Engine", icon="🪟", default=True)
frame_page = st.Page("pages/1_Frame_Details.py", title="Frame Size (WxH) & SQFT", icon="🖼️")

pg = st.navigation({"Main Menu": [main_page], "Tools": [frame_page]})
pg.run()
