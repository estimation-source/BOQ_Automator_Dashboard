import streamlit as st
import pandas as pd
import io
import re

# 🔗 main.py मधून फंक्शन आणि डेटा स्ट्रक्चर Import करत आहोत
from main import process_uploaded_files, get_image_path

# ============================================================
# Page Configuration
# ============================================================
st.set_page_config(
    page_title="WIN-SQUARE | Frame Size & Clean Summary",
    layout="wide",
    page_icon="🖼️"
)

st.title("🖼️ Window Frame Size & Clean Summary Engine")
st.caption("फक्त Frame Width x Height, Frame SQFT आणि Clean Glass Type Summary पाहण्यासाठी स्वतंत्र ॲप")

st.markdown("---")

# ============================================================
# Step 1: File Upload Section
# ============================================================
st.subheader("📁 Step 1: Upload BOQ Excel Files")

uploaded_files = st.file_uploader(
    "Upload BOQ Excel Files",
    type=["xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
    key="frame_app_uploader"
)

if st.button("🚀 Frame Data & Clean Summary Generate करा", type="primary"):
    if uploaded_files:
        with st.spinner("Processing Frame & Glass Data..."):
            # main.py मधील प्रोसेसिंग फंक्शन कॉल केले
            df_merged = process_uploaded_files(uploaded_files)

            if not df_merged.empty:
                st.session_state["frame_df"] = df_merged
                st.success("✅ डेटा यशस्वीरित्या एक्सट्रॅक्ट झाला आहे!")
            else:
                st.error("⚠️ फाईलमध्ये कोणताही वैध डेटा सापडला नाही.")
    else:
        st.warning("कृपया आधी Excel फाईल अपलोड करा.")

# ============================================================
# Step 2: Clean Data & Frame Calculations
# ============================================================
if "frame_df" in st.session_state:
    df = st.session_state["frame_df"].copy()

    # 1. Frame Calculations (Width x Height -> SQFT)
    df["Frame WxH"] = df["Width"].astype(str) + " x " + df["Height"].astype(str)
    df["Single Frame SQFT"] = ((df["Width"] * df["Height"]) / 92903.04).round(4)
    df["Total Frame SQFT"] = (df["Single Frame SQFT"] * df["Qty"]).round(4)

    # 2. Clean Glass Type Cleaning Function
    def clean_glass_name(text):
        if pd.isna(text) or not str(text).strip():
            return "NOT SPECIFIED"
        text = str(text).upper().strip()
        text = re.sub(r"\s+", " ", text)
        text = text.replace("TOUGHENED", "THGN").replace("TOUGHNED", "THGN")
        return text

    df["Clean Glass Spec"] = df["GlassType"].apply(clean_glass_name)

    # KPI Summary Cards
    st.markdown("### 📊 Frame Quick Overview")
    k1, k2, k3 = st.columns(3)
    k1.metric("Total Windows / Frames", len(df))
    k2.metric("Total Quantity (Pcs)", int(df["Qty"].sum()))
    k3.metric("Total Frame Area (SQFT)", f"{df['Total Frame SQFT'].sum():,.2f} Sq.Ft")

    st.markdown("---")

    # Tabs for Clean Data View
    tab1, tab2 = st.tabs(["🪟 Window Frame Size Details", "📊 OC & Clean Glass Type Summary"])

    # --------------------------------------------------------
    # TAB 1: WINDOW FRAME WxH & SQFT DETAILS
    # --------------------------------------------------------
    with tab1:
        st.subheader("🖼️ Frame Dimensions and Area Details")
        
        display_frame_df = df[[
            "WindowCode", "Width", "Height", "Frame WxH", 
            "Qty", "Single Frame SQFT", "Total Frame SQFT", 
            "Clean Glass Spec", "SourceFile"
        ]].copy()

        display_frame_df.columns = [
            "Window Code", "Width (mm)", "Height (mm)", "Frame Size (WxH)", 
            "Qty", "SQFT / Pc", "Total Frame SQFT", 
            "Glass Spec", "Source File"
        ]

        display_frame_df.insert(0, "Sr. No.", range(1, len(display_frame_df) + 1))
        st.dataframe(display_frame_df, use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # TAB 2: OC WISE & CLEAN GLASS TYPE SUMMARY
    # --------------------------------------------------------
    with tab2:
        st.subheader("📊 OC Wise Clean Glass Summary")

        def make_glass_summary_str(group):
            summary = group.groupby("Clean Glass Spec")["Qty"].sum()
            details = [f"{g_type}: {qty} Pcs" for g_type, qty in summary.items() if g_type != "NOT SPECIFIED"]
            return " | ".join(details) if details else "Not Specified"

        oc_glass_summary = df.groupby("SourceFile").apply(make_glass_summary_str).reset_index(name="Glass Type Breakdown")

        oc_totals = df.groupby("SourceFile", as_index=False).agg(
            Total_Qty=("Qty", "sum"),
            Total_Frame_SQFT=("Total Frame SQFT", "sum")
        )

        oc_clean_summary = pd.merge(oc_totals, oc_glass_summary, on="SourceFile")
        oc_clean_summary["Total_Frame_SQFT"] = oc_clean_summary["Total_Frame_SQFT"].round(2)
        
        oc_clean_summary.columns = ["Source File (OC Name)", "Total Qty (Pcs)", "Total Frame SQFT", "Glass Breakdown (Clean)"]
        oc_clean_summary.insert(0, "Sr. No.", range(1, len(oc_clean_summary) + 1))

        st.dataframe(oc_clean_summary, use_container_width=True, hide_index=True)
