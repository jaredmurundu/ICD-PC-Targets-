# ICD Performance Contract (PC) Tracking Dashboard

This is a Streamlit web application for tracking performance contract (PC) targets at the Institute of Co-operative Development (ICD), Cooperative University of Kenya.

## 📊 Features

- View and edit PC tracking matrix interactively
- Assign responsible officers and update progress
- Save changes directly to Excel
- Download updated matrix
- User-friendly, responsive interface

## 🚀 How to Deploy on Streamlit Community Cloud

1. **Fork or clone this repository**  
   Make sure your repo includes:
   - `PC TRACKING ACTIVITIES.py`
   - `PC_Tracking_Matrix_ICD_2024_2025.xlsx`
   - `requirements.txt`

2. **Go to [https://share.streamlit.io](https://share.streamlit.io)**  
   - Log in with your GitHub account
   - Select this repo
   - Set `PC TRACKING ACTIVITIES.py` as the app file
   - Click Deploy

3. **That's it!**  
   Your app will be live at `https://yourname.streamlit.app/`

## ⚙️ Local Setup

If you prefer to run the app locally:

```bash
git clone https://github.com/yourname/icd-pc-tracking.git
cd icd-pc-tracking
pip install -r requirements.txt
streamlit run "PC TRACKING ACTIVITIES.py"
```

## 📝 Notes

- Make sure `PC_Tracking_Matrix_ICD_2024_2025.xlsx` is in the same directory as your app script.
- Edits are saved locally — Streamlit Community Cloud resets files after each session.

## 📫 Contact

For questions or improvements, please contact:  
**Jared Murundu**  
📧 jmurundu@cuk.ac.ke

---

Developed for internal performance monitoring use at ICD - The Co-operative University of Kenya.