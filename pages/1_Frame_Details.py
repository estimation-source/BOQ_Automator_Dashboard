import io
import re
import math
import openpyxl
import pandas as pd
import streamlit as st

# Page Config
st.set_page_config(
    page_title="WIN-SQUARE | Frame WxH Details",
    layout="wide",
    page_icon="🖼️"
)

st.title("🖼️ Window Frame Size (WxH) & Area Engine")
st.caption("Client Details Sheet च्या पुढील sheet मधून Frame Width x Height, Qty आणि Total SQFT मोजणारा स्वतंत्र रिपोर्ट")
st.markdown("---")

def extract_frame_data_from_workbook(file) -> pd.DataFrame:
    file_bytes = io.BytesIO(file.read())
    file.seek(0)
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    
    records = []
    all_sheets = wb.sheetnames
    target_sheet_name = None
    
    # Client Details नंतर येणारी पहिली Sheet शोधणे
    client_idx = -1
    for i, s_name in enumerate(all_sheets):
        if any(k in s_name.upper() for k in ["CLIENT", "DETAIL", "CUSTOMER"]):
            client_idx = i
            break
            
    if client_idx != -1 and client_idx + 1 < len(all_sheets):
        target_sheet_name = all_sheets[client_idx + 1]
    else:
        target_sheet_name = all_sheets[0] if len(all_sheets) == 1 else all_sheets[1]

    sheet = wb[target_sheet_name]
    data_rows = list(sheet.values)
    if not data_rows:
        return pd.DataFrame()

    df = pd.DataFrame(data_rows)
    
    # Columns Index शोधणे
    f_w_col, f_h_col, qty_col, code_col = None, None, None, None
    
    for r_idx in range(min(15, len(df))):
        row = [str(val).upper().strip() if val is not None else "" for val in df.iloc[r_idx]]
        for c_idx, val in enumerate(row):
            if any(k in val for k in ["FRAME W", "FRAME WIDTH", "FRAME_W", "FWIDTH"]):
                f_w_col = c_idx
            if any(k in val for k in ["FRAME H", "FRAME HEIGHT", "FRAME_H", "FHEIGHT"]):
                f_h_col = c_idx
            if "QTY" in val or "QUANTITY" in val:
                qty_col = c_idx
            if "CODE" in val or "WINDOW" in val:
                code_col = c_idx

    # जर स्पेसिफिक Frame Column नाव सापडले नाही तर बेसिक W आणि H शोधणे
    if f_w_col is None or f_h_col is None:
        for r_idx in range(min(15, len(df))):
            row = [str(val).upper().strip() if val is not None else "" for val in df.iloc[r_idx]]
            for c_idx, val in enumerate(row):
                if val in ["WIDTH", "W", "W(MM)"]:
                    f_w_col = c_idx
                if val in ["HEIGHT", "H", "H(MM)"]:
                    f_h_col = c_idx

    if f_w_col is None or f_h_col is None:
        return pd.DataFrame()

    # डेटा रीड करणे
    for r_idx in range(1, len(df)):
        row = df.iloc[r_idx]
        try:
            w_val = float(row[f_w_col]) if pd.notna(row[f_w_col]) else None
            h_val = float(row[f_h_col]) if pd.notna(row[f_h_col]) else None
            
            if w_val and h_val and w_val > 0 and h_val > 0:
                qty = 1
                if qty_col is not None and pd.notna(row[qty_col]):
                    try:
                        qty = int(float(row[qty_col]))
                    except:
                        qty = 1
                        
                code = f"WIN-{r_idx}"
                if code_col is not None and pd.notna(row[code_col]):
                    code = str(row[code_col]).strip()

                single_sqft = round((w_val * h_val) / 92903.04, 4)
                total_sqft = round(single_sqft * qty, 4)

                records.append({
                    "Window Code": code,
                    "Frame Width (mm)": int(w_val),
                    "Frame Height (mm)": int(h_val),
                    "Frame Size (WxH)": f"{int(w_val)} x {int(h_val)}",
                    "Qty (Pcs)": qty,
                    "Single Frame SQFT": single_sqft,
                    "Total Frame SQFT": total_sqft,
                    "Extracted Sheet": target_sheet_name,
                    "Source File": file.name
                })
        except:
            continue

    return pd.DataFrame(records)


# Streamlit UI Setup
uploaded_files = st.file_uploader(
    "Upload BOQ Excel Files", 
    type=["xlsx", "xlsm"], 
    accept_multiple_files=True,
    key="frame_uploader_page"
)

if st.button("🚀 Process Frame Data", type="primary"):
    if uploaded_files:
        all_data = []
        with st.spinner("Processing Client Details Next Sheet..."):
            for file in uploaded_files:
                res_df = extract_frame_data_from_workbook(file)
                if not res_df.empty:
                    all_data.append(res_df)
            
            if all_data:
                st.session_state["frame_page_df"] = pd.concat(all_data, ignore_index=True)
                st.success("✅ Frame Data successfully extracted!")
            else:
                st.error("⚠️ फाईलमध्ये Client Details च्या नंतरच्या sheet वर Frame Size चे columns सापडले नाहीत.")
    else:
        st.warning("कृपया Excel फाईल्स अपलोड करा.")

if "frame_page_df" in st.session_state:
    final_df = st.session_state["frame_page_df"]
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # KPI Metrics Cards
    k1, k2, k3 = st.columns(3)
    k1.metric("Total Windows Count", len(final_df))
    k2.metric("Total Frame Qty", int(final_df["Qty (Pcs)"].sum()))
    k3.metric("GRAND TOTAL FRAME SQFT", f"{final_df['Total Frame SQFT'].sum():,.2f} Sq.Ft")
    
    st.markdown("---")
    
    t1, t2 = st.tabs(["📋 Itemized Frame Details", "📊 File-Wise Frame Summary"])
    
    with t1:
        st.dataframe(final_df, use_container_width=True, hide_index=True)
        
    with t2:
        summary_df = final_df.groupby("Source File", as_index=False).agg(
            Total_Qty=("Qty (Pcs)", "sum"),
            Total_Frame_SQFT=("Total Frame SQFT", "sum")
        )
        summary_df["Total_Frame_SQFT"] = summary_df["Total_Frame_SQFT"].round(2)
        summary_df.columns = ["Source File (OC Name)", "Total Qty (Pcs)", "Total Frame SQFT"]
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
