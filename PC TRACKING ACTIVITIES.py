import streamlit as st
import pandas as pd
from io import BytesIO



st.set_page_config(page_title="ICD PC Tracking Dashboard", layout="wide")

# University Header
st.markdown("""
    <div style='text-align: center; margin-bottom: 20px; font-family: Garamond, serif;'>
        <img src='https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Cuk.png' width='160' />
        <p style='font-size: 18px;'><strong>THE CO-OPERATIVE UNIVERSITY OF KENYA</strong></p>
        <p>P.O. Box 24814 – 00502, Karen, Kenya</p>
        <p>Telephone: (020)-2430127 / 2679456 / 8891401 &nbsp;&nbsp; Fax: (020)-8891410</p>
        <p>Website: <a href='https://www.cuk.ac.ke' target='_blank'>www.cuk.ac.ke</a> &nbsp;&nbsp; Email: enquiries@cuk.ac.ke</p>
        <p style='font-size: 16px; margin-top: 10px;'><strong>DIVISION OF ACADEMICS, CO-OPERATIVE DEVELOPMENT, RESEARCH AND INNOVATION (ACDRI)</strong></p>
        <p><strong>INSTITUTE OF CO-OPERATIVE DEVELOPMENT (ICD)</strong></p>
    </div>
""", unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# Centered, uppercase title in Garamond
st.markdown("""
    <h3 style='text-align: center; text-transform: uppercase; font-family: Garamond, serif; margin-bottom: 10px;'>
        Vice Chancellor – ACDRI
    </h3>
""", unsafe_allow_html=True)

# Profile section for Prof. Isaac Nyamongo
st.markdown("""
    <div style='text-align: center; font-family: Garamond, serif; margin-bottom: 30px;'>
        <img src='https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Prof%20Isaac%20Nyamongo.jpg' width='160'/>
        <p style='font-size: 16px; margin-top: 10px;'><strong>Prof. Isaac Nyamongo</strong><br>Deputy Vice Chancellor – ACDRI</p>
    </div>
""", unsafe_allow_html=True)


# Section Header
st.markdown("### ICD STAFF")

# Create 4 side-by-side columns for staff
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.image("https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Prof. Wycliffe Oboka.webp", width=130)
    st.markdown("**Prof. Wycliffe Oboka**  \nDirector, ICD", unsafe_allow_html=True)

with col2:
    st.image("https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Victor Wambua.png", width=130)
    st.markdown("**Victor Wambua**  \nShort Courses Co-ordinator", unsafe_allow_html=True)

with col3:
    st.image("https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Jared Murundu.jpg", width=130)
    st.markdown("**Jared Murundu**  \nAdministrative Assistant", unsafe_allow_html=True)

with col4:
    st.image("https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Mercy Sipayo.jpg", width=130)
    st.markdown("**Mercy Sipayo**  \nOffice Assistant (Records)", unsafe_allow_html=True)


st.markdown("""
    <h1 style='text-align: center; text-transform: uppercase; font-family: Garamond, serif;'>
         ICD PERFORMANCE CONTRACT (PC) TRACKING DASHBOARD
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
            "The PC Target": [],
            "Responsible ICD Officer": [],
            "Status of The PC Target": [],
            "Evidence": []
        })

# Load the data
df = load_data()

st.subheader("Please Update PC Matrix Below")

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


