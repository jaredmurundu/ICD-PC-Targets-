import streamlit as st
import pandas as pd
from io import BytesIO

# Title
st.set_page_config(page_title="ICD PC Tracking Dashboard", layout="wide")
st.title("📊 ICD Performance Contract (PC) Tracking Dashboard")

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

# Save Updates
if st.button("💾 Save Updates"):
    edited_df.to_excel("PC_Tracking_Matrix_ICD_2024_2025.xlsx", index=False)
    st.success("✅ Updates saved successfully!")

# Download Button
st.subheader("⬇️ Download Updated PC Matrix")
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='PC_Tracking')
    processed_data = output.getvalue()
    return processed_data

# Inject custom CSS styling
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

# Render the actual download button
st.download_button(
    label="Download Excel File",
    data=to_excel(edited_df),
    file_name='Updated_PC_Tracking_Matrix.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

)

st.markdown("""
<hr>
<div style='text-align: center; font-family: Garamond, serif; font-size: 18px;'>
    <strong>Developed by Jared Murundu for ACDRI (ICD/DRI) © All Rights Reserved | CUK • 2025</strong>
</div>
""", unsafe_allow_html=True)


