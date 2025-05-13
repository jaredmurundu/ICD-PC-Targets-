import streamlit as st
import pandas as pd
from io import BytesIO



st.set_page_config(page_title="ICD PC Tracking Dashboard", layout="wide")

st.markdown("""
    <div style='text-align: center; margin-bottom: 10px;'>
        <img src='https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Cuk.png' width='160'/>
        <p style='font-family: Garamond, serif; font-size: 16px;'>📧 enquiries@cuk.ac.ke| +254 724311606</p>
    </div>
""", unsafe_allow_html=True)

st.markdown("""
    <h1 style='text-align: center; text-transform: uppercase; font-family: Garamond, serif;'>
        📊 ICD PERFORMANCE CONTRACT (PC) TRACKING DASHBOARD
    </h1>
""", unsafe_allow_html=True)


# Load or create DataFrame
@st.cache_data

def load_data():
    file_path = "PC_Tracking_Matrix_ICD_2024_2025.xlsx"
    try:
        return pd.read_excel(file_path)
    except FileNotFoundError:
        return pd.DataFrame({
            "PC Target": [],
            "Responsible Officer": [],
            "Status of Achievement": [],
            "Evidence": []
        })

# Load the data
df = load_data()

st.subheader("📋 Please Edit or Update PC Matrix Below")

edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

st.markdown("""
    <style>
    div.stButton > button {
        background-color: green !important;
        color: white !important;
        font-weight: bold;
        border-radius: 6px;
    }
    </style>
""", unsafe_allow_html=True)

# Save button
if st.button("💾 Save"):
    edited_df.to_excel("PC_Tracking_Matrix_ICD_2024_2025.xlsx", index=False)
    st.success("✅ Saved!")

# Download Button
st.subheader("⬇️ Download Updated PC Matrix")
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='PC_Tracking')
    processed_data = output.getvalue()
    return processed_data

st.markdown("""
    <style>
    .stDownloadButton>button {
        background-color: yellow !important;
        color: blue !important;
        border: 2px solid green !important;
        font-weight: bold;
        font-family: Garamond, serif;
        padding: 10px 20px;
        border-radius: 10px;
    }
    </style>
""", unsafe_allow_html=True)

st.download_button(
    label="Download Excel File",
    data=to_excel(edited_df),
    file_name='Updated_PC_Tracking_Matrix.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)


st.markdown("""
<hr>
<div style='text-align: center; font-family: Garamond, serif; font-size: 18px;'>
    <strong>Developed by Jared Murundu for ACDRI (ICD/DRI) © All Rights Reserved | CUK • 2025</strong>
</div>
""", unsafe_allow_html=True)


