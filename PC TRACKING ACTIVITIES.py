import streamlit as st
import pandas as pd
from io import BytesIO



st.set_page_config(page_title="ICD PC Tracking Dashboard", layout="wide")

st.markdown("""
    <div style='text-align: center; margin-bottom: 20px; font-family: Garamond, serif;'>
        <img src='https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Cuk.png' width='160' />
        <p style='font-size: 18px; margin: 4px 0;'><strong>THE CO-OPERATIVE UNIVERSITY OF KENYA</strong></p>
        <p style='font-size: 16px; margin: 2px 0;'>P.O. Box 24814 – 00502, Karen, Kenya</p>
        <p style='font-size: 16px; margin: 2px 0;'>Telephone: (020)-2430127 / 2679456 / 8891401 &nbsp;&nbsp; Fax: (020)-8891410</p>
        <p style='font-size: 16px; margin: 2px 0;'>Website: <a href='https://www.cuk.ac.ke' target='_blank'>www.cuk.ac.ke</a> &nbsp;&nbsp; Email: enquiries@cuk.ac.ke</p>
        <p style='font-size: 16px; margin: 10px 0;'><strong>DIVISION OF ACADEMICS, CO-OPERATIVE DEVELOPMENT, RESEARCH AND INNOVATION(ACDRI)</strong></p>
    </div>
""", unsafe_allow_html=True)
st.markdown("""
    <div style='text-align: center; font-family: Garamond, serif; margin-top: 20px;'>

        <h3 style='margin-bottom: 5px;'>DIVISION OF ACADEMICS, CO-OPERATIVE DEVELOPMENT, RESEARCH AND INNOVATION (ACDRI)</h3>
        <h4 style='margin-top: 0px;'>INSTITUTE OF CO-OPERATIVE DEVELOPMENT (ICD)</h4>

        <div style='display: flex; justify-content: center; gap: 40px; flex-wrap: wrap; margin-top: 20px;'>

            <!-- Director -->
            <div style='text-align: center; width: 200px;'>
                <img src='https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/director.png' width='120'/>
                <p style='font-size: 16px; margin-top: 5px;'><strong>Prof. Wycliffe Oboka</strong><br/>Director, ICD</p>
            </div>

            <!-- Short Courses Coordinator -->
            <div style='text-align: center; width: 200px;'>
                <img src='https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/victor.png' width='120'/>
                <p style='font-size: 16px; margin-top: 5px;'><strong>Victor Wambua</strong><br/>Short Courses Co-ordinator</p>
            </div>

            <!-- Admin Assistant -->
            <div style='text-align: center; width: 200px;'>
                <img src='https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/jared.png' width='120'/>
                <p style='font-size: 16px; margin-top: 5px;'><strong>Jared Murundu</strong><br/>Administrative Assistant</p>
            </div>

            <!-- Office Assistant -->
            <div style='text-align: center; width: 200px;'>
                <img src='https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/mercy.png' width='120'/>
                <p style='font-size: 16px; margin-top: 5px;'><strong>Mercy Sipayo</strong><br/>Office Assistant (Records)</p>
            </div>

        </div>
    </div>
""", unsafe_allow_html=True)

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


