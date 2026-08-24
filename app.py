import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from streamlit_geolocation import streamlit_geolocation
from streamlit_folium import st_folium
import folium
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Village Coverage 2026 Form", layout="centered")

st.title("📍 Village Coverage 2026 Form")

# ----------------- GOOGLE SHEET CONNECTION -----------------
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    if "gcp_service_account" not in st.secrets:
        raise Exception("Streamlit Secrets me 'gcp_service_account' nahi mila.")
        
    creds_dict = dict(st.secrets["gcp_service_account"])
    
    # Private Key formatting fix (Fixes MalformedFraming / PEM Error)
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(credentials)

SPREADSHEET_ID = "1WptCID2zXSEqUvWbCam23HstE2W64RYj4bBRZsdefsA"

try:
    gc = get_gspread_client()
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    sheet_master = spreadsheet.worksheet("RD To Spoke Data")
    # Screenshot ke hisab se exact tab name fix kiya hai:
    sheet_survey = spreadsheet.worksheet("2nd Village Coverage 2026") 
except Exception as e:
    st.error(f"❌ Connection Error: {type(e).__name__} - {str(e)}")
    st.stop()

# Master Data Read Function
@st.cache_data(ttl=60)
def load_data():
    data = sheet_master.get_all_records()
    df_rd = pd.DataFrame(data)
    df_rd.columns = df_rd.columns.str.strip()
    for col in df_rd.select_dtypes(include=['object']).columns:
        df_rd[col] = df_rd[col].astype(str).str.strip()
    return df_rd

try:
    df_master = load_data()
except Exception as e:
    st.error(f"❌ Error loading Master Data: {e}")
    st.stop()

# ----------------- SESSION STATE SETUP -----------------
if 'form_counter' not in st.session_state:
    st.session_state.form_counter = 0

fc = st.session_state.form_counter

if f'unique_id_{fc}' not in st.session_state:
    st.session_state[f'unique_id_{fc}'] = f"UID_{int(datetime.now().timestamp() * 1000)}"

current_uid = st.session_state[f'unique_id_{fc}']

st.subheader("Survey Details")
st.info(f"🆔 Session Unique ID: **{current_uid}**")

ist_time = datetime.utcnow() + timedelta(hours=5, minutes=30)

col1, col2 = st.columns(2)
with col1:
    current_date = ist_time.strftime("%Y-%m-%d")
    st.text_input("Date", value=current_date, disabled=True, key=f"display_date_{fc}")
with col2:
    current_time = ist_time.strftime("%H:%M:%S")
    st.text_input("Time", value=current_time, disabled=True, key=f"display_time_{fc}")

# ----------------- DROPDOWNS (CASCADE FILTERING) -----------------
# 1. RD Name
rd_options = ["Select..."] + sorted([x for x in df_master['RD NAME'].unique() if x and x != 'nan'])
selected_rd = st.selectbox("RD Name *", rd_options, key=f"rd_name_{fc}")

df_f1 = df_master[df_master['RD NAME'] == selected_rd] if selected_rd != "Select..." else pd.DataFrame(columns=df_master.columns)

# 2. S.E Name
se_options = ["Select..."] + sorted([x for x in df_f1['S.E Name'].unique() if x and x != 'nan']) if not df_f1.empty else ["Select..."]
selected_se = st.selectbox("STL / S.E Name *", se_options, key=f"se_name_{fc}")

df_f2 = df_f1[df_f1['S.E Name'] == selected_se] if selected_se != "Select..." and not df_f1.empty else pd.DataFrame(columns=df_master.columns)

# 3. ASM Name
asm_options = ["Select..."] + sorted([x for x in df_f2['Asm Name'].unique() if x and x != 'nan']) if not df_f2.empty else ["Select..."]
selected_asm = st.selectbox("ASM Name *", asm_options, key=f"asm_name_{fc}")

df_f3 = df_f2[df_f2['Asm Name'] == selected_asm] if selected_asm != "Select..." and not df_f2.empty else pd.DataFrame(columns=df_master.columns)

# 4. SM Name
sm_options = ["Select..."] + sorted([x for x in df_f3['Sm Name'].unique() if x and x != 'nan']) if not df_f3.empty else ["Select..."]
selected_sm = st.selectbox("SM Name *", sm_options, key=f"sm_name_{fc}")

df_f4 = df_f3[df_f3['Sm Name'] == selected_sm] if selected_sm != "Select..." and not df_f4.empty else pd.DataFrame(columns=df_master.columns)

# 5. Distributor Name & Code
dist_options = ["Select..."] + sorted([x for x in df_f4['Distributor Name, Town DRB Code'].dropna().unique() if x and x != 'nan']) if not df_f4.empty else ["Select..."]
selected_dist = st.selectbox("Distributor Name & Code *", dist_options, key=f"dist_name_{fc}")

df_f5 = df_f4[df_f4['Distributor Name, Town DRB Code'] == selected_dist] if selected_dist != "Select..." and not df_f4.empty else pd.DataFrame(columns=df_master.columns)

# 6. Spoke Name & Code
spoke_options = ["Select..."] + sorted([x for x in df_f5['Spoke Name, Town Spoke Code'].dropna().unique() if x and x != 'nan']) if not df_f5.empty else ["Select..."]
selected_spoke = st.selectbox("Spoke Name & Code *", spoke_options, key=f"spoke_name_{fc}")

# 7. Village Name
entered_village = st.text_input("Village Name * (Type here)", key=f"village_name_{fc}")

# 8. Covered / Uncovered
coverage_status = st.selectbox("Covered / Uncovered *", ["Select...", "Covered", "Uncovered"], key=f"coverage_status_{fc}")

# 9. Outlet In Village
outlet_count = st.number_input("Outlet In Village", min_value=0, value=0, step=1, key=f"outlet_count_{fc}")

# ----------------- LOCATION CAPTURE -----------------
st.markdown("---")
st.subheader("🌐 Location Capture & OpenStreetMap")
st.write("Click below to capture GPS location:")

loc = streamlit_geolocation()

if loc and loc.get('latitude') and loc.get('longitude'):
    lat = loc['latitude']
    lon = loc['longitude']
    acc = loc.get('accuracy', 0)
    
    if acc > 30:
        st.warning(f"⚠️ Warning: GPS Accuracy is poor ({acc:.1f} meters). Please move to an open area.")
    else:
        st.success(f"📍 Excellent GPS Accuracy! ({acc:.1f} meters)")
    
    m = folium.Map(location=[lat, lon], zoom_start=17, tiles="OpenStreetMap")
    folium.Marker(
        [lat, lon],
        popup=f"<b>{entered_village if entered_village else 'Survey Location'}</b><br>Accuracy: {acc:.1f}m",
        tooltip="Captured Location",
        icon=folium.Icon(color="blue" if acc <= 30 else "orange", icon="info-sign")
    ).add_to(m)
    
    st_folium(m, width=700, height=400, key=f"map_{fc}")

# ----------------- DATA APPEND TO GOOGLE SHEET -----------------
if 'submitted_successfully' not in st.session_state:
    st.session_state.submitted_successfully = False

if st.session_state.submitted_successfully:
    st.success("🎉 Form Successfully Saved to Google Sheet!")
    if st.button("➕ Fill Next Form", type="primary", key=f"next_btn_{fc}"):
        st.session_state.submitted_successfully = False
        st.session_state.form_counter += 1
        st.rerun()
else:
    if st.button("Save Form", type="primary", key=f"save_btn_{fc}"):
        if selected_rd == "Select..." or selected_se == "Select..." or selected_dist == "Select..." or selected_spoke == "Select..." or not entered_village.strip() or coverage_status == "Select...":
            st.error("❌ Kripya sabhi zaroori fields (* marked) bharein!")
        elif not loc or not loc.get('latitude'):
            st.error("❌ Location capture nahi hui! Kripya GPS allow karein.")
        else:
            lat = loc.get('latitude')
            lon = loc.get('longitude')
            location_str = f"{lat}, {lon}"
            
            sub_date = ist_time.strftime("%Y-%m-%d")
            sub_time = ist_time.strftime("%H:%M:%S")
            
            row_data = [
                current_uid,
                sub_date,
                sub_time,
                selected_rd,
                selected_se,
                selected_asm if selected_asm != "Select..." else "",
                selected_sm if selected_sm != "Select..." else "",
                entered_village.strip(),
                coverage_status,
                selected_dist,
                selected_spoke,
                int(outlet_count),
                location_str
            ]
            
            try:
                sheet_survey.append_row(row_data, value_input_option='USER_ENTERED')
                st.session_state.submitted_successfully = True
                st.rerun()
            except Exception as ex:
                st.error(f"❌ Error saving form to Google Sheet: {ex}")

# ----------------- ADMIN PANEL -----------------
st.sidebar.markdown("---")
st.sidebar.subheader("🔒 Admin Download Panel")
admin_password = st.sidebar.text_input("Enter Password to Download", type="password")

if admin_password == "slmg2026":
    st.sidebar.success("✅ Access Granted")
    
    try:
        survey_records = sheet_survey.get_all_records()
        df_download = pd.DataFrame(survey_records)
        
        csv_data = df_download.to_csv(index=False).encode('utf-8')
        
        st.sidebar.download_button(
            label="📥 Download Full Survey Data (CSV)",
            data=csv_data,
            file_name="Village_Coverage_Survey_2026.csv",
            mime="text/csv"
        )
    except Exception as e:
        st.sidebar.error(f"Error preparing download: {e}")
elif admin_password != "":
    st.sidebar.error("❌ Incorrect Password")
