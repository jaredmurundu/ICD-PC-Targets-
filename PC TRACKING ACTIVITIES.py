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
        <p style='font-size: 16px; margin: 10px 0;'><strong>DIVISION OF ACADEMICS, CO-OPERATIVE DEVELOPMENT, RESEARCH AND INNOVATION</strong></p>
    </div>
""", unsafe_allow_html=True)


# Section divider
st.markdown("<hr>", unsafe_allow_html=True)

# Vice Chancellor Section
st.markdown("""
    <div style='text-align: center; font-family: Garamond, serif; margin-bottom: 40px;'>
        <h3 style='text-transform: uppercase; margin-bottom: 10px;'>Vice Chancellor – ACDRI</h3>
        <div style='display: inline-block; padding: 15px; border-radius: 15px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2); transition: transform 0.3s;'>
            <img src='https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Prof%20Isaac%20Nyamongo.jpeg' width='160' style='border-radius: 10px;'/>
            <p style='font-size: 17px; margin-top: 12px;'><strong>Prof. Isaac Nyamongo</strong><br>Deputy Vice Chancellor – ACDRI</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# Section Heading
st.markdown("<h3 style='text-align: center; font-family: Garamond, serif;'>ICD STAFF </h3>", unsafe_allow_html=True)

# ICD Team Members
col1, col2, col3, col4 = st.columns(4)

def render_card(image_url, name, title):
    st.markdown(f"""
        <div style='text-align: center; font-family: Garamond, serif; 
                    border-radius: 15px; padding: 15px; 
                    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2); 
                    transition: transform 0.3s ease;'>
            <img src='{image_url}' width='130' style='border-radius: 10px; margin-bottom: 10px;'/>
            <p style='font-size: 16px;'><strong>{name}</strong><br>{title}</p>
        </div>
    """, unsafe_allow_html=True)

with col1:
    render_card(
        "https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Prof.%20Wycliffe%20Oboka.webp",
        "Prof. Wycliffe Oboka",
        "Director, ICD"
    )

with col2:
    render_card(
        "https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Victor%20Wambua.png",
        "Mr. Victor Wambua",
        "Short Courses Co-ordinator, ICD"
    )

with col3:
    render_card(
        "https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Jared%20Murundu.jpg",
        "Mr. Jared Murundu",
        "Administrative Assistant, ICD"
    )

with col4:
    render_card(
        "https://raw.githubusercontent.com/jaredmurundu/ICD-PC-Targets-/main/Mercy%20Sipayo.jpg",
        "Ms Mercy Sipayo",
        "Office Assistant (Records), ICD"
    )


@st.cache_data
def load_data():
    file_path = "PC_Tracking_Matrix_ICD_2024_2025.xlsx"
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        # Start with placeholder text to enforce string type
        df = pd.DataFrame({
            "The PC Target": ["Example Target"],
            "Responsible ICD Officer": ["Enter Name"],
            "Status of The PC Target": ["Pending / Done / Ongoing"],
            "Evidence": ["Supporting docs, links"]
        })

    # Force all columns to be treated as strings
    for col in df.columns:
        df[col] = df[col].astype(str).fillna("")

    return df

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

if st.button("Save"):
    edited_df.to_excel("PC_Tracking_Matrix_ICD_2024_2025.xlsx", index=False)
    st.success("✅ The Changes are Saved, Thank you.!")

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
# ICD External Website Link
st.markdown("""
<hr>
<div style='text-align: center; font-family: Garamond, serif; font-size: 16px; margin-top: 30px;'>
    📎 <a href='https://cuk.ac.ke/icd/' target='_blank'><strong>Visit the ICD Website</strong></a>
</div>
""", unsafe_allow_html=True)

# Brochure download/view
st.markdown("""
<hr>
<h4 style='text-align: center; font-family: Garamond, serif;'>📘 ICD Brochure</h4>

<div style='text-align: center;'>
    <a href='https://github.com/jaredmurundu/ICD-PC-Targets-/raw/main/ICD-Brochure.pdf' target='_blank'>
        <button style='background-color: #4CAF50; color: white; font-size: 16px; padding: 10px 20px; border: none; border-radius: 8px;'>
            📥 View / Download ICD Brochure (PDF)
        </button>
    </a>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<hr>
<div style='text-align: center; font-family: Garamond, serif; font-size: 18px;'>
    <strong>Developed by Jared Murundu for ACDRI (ICD/DRI) © All Rights Reserved | CUK • 2025</strong>
</div>
""", unsafe_allow_html=True)


