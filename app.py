import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from streamlit_geolocation import streamlit_geolocation
from streamlit_folium import st_folium
import folium
import os
import json
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Village Coverage 2026 Form", layout="centered")
st.title("📍 Village Coverage 2026 Form")

file_path = "Village 2026.xlsx"
backup_dir = "survey_backups"

if not os.path.exists(backup_dir):
    os.makedirs(backup_dir)

def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        # Fix line breaks in private key for Streamlit secrets
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    elif os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file("service_account.json", scopes=scopes)
    else:
        st.error("❌ Service account credentials missing!")
        st.stop()
        
    return gspread.authorize(creds)

def sync_to_google_sheet(record):
    try:
        gc = get_gspread_client()
        sheet = gc.open("Village Coverage 2026").worksheet("Village Coverage 2026")
        
        row_data = [
            record.get('UNIQUE_ID', ''),
            record.get('Date', ''),
            record.get('Time', ''),
            record.get('RD Name', ''),
            record.get('STL NAMe', ''),
            record.get('ASM Name', ''),
            record.get('SM Name', ''),
            record.get('Village Name', ''),
            record.get('Covered/Uncovered', ''),
            record.get('Distributor Name & Code', ''),
            record.get('Spoke Name & Code', ''),
            record.get('Outlet In Village', 0),
            record.get('Location', '')
        ]
        
        sheet.append_row(row_data)
        return True, "Synced to Google Sheet"
    except Exception as e:
        return False, str(e)

@st.cache_data(ttl=1)
def load_data():
    xls = pd.ExcelFile(file_path)
    df_rd = pd.read_excel(xls, sheet_name="RD To Spoke Data")
    df_rd.columns = df_rd.columns.str.strip()
    for col in df_rd.select_dtypes(include=['object']).columns:
        df_rd[col] = df_rd[col].astype(str).str.strip()
    return df_rd

try:
    df_master = load_data()
except Exception as e:
    st.error(f"❌ Error loading Excel file: {e}")
    st.stop()

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

# Dropdowns Setup
rd_options = ["Select..."] + sorted([x for x in df_master['RD NAME'].unique() if x and x != 'nan'])
selected_rd = st.selectbox("RD Name *", rd_options, key=f"rd_name_{fc}")

df_f1 = df_master[df_master['RD NAME'] == selected_rd] if selected_rd != "Select..." else pd.DataFrame(columns=df_master.columns)

se_options = ["Select..."] + sorted([x for x in df_f1['S.E Name'].unique() if x and x != 'nan']) if not df_f1.empty else ["Select..."]
selected_se = st.selectbox("STL / S.E Name *", se_options, key=f"se_name_{fc}")

df_f2 = df_f1[df_f1['S.E Name'] == selected_se] if selected_se != "Select..." and not df_f1.empty else pd.DataFrame(columns=df_master.columns)

asm_options = ["Select..."] + sorted([x for x in df_f2['Asm Name'].unique() if x and x != 'nan']) if not df_f2.empty else ["Select..."]
selected_asm = st.selectbox("ASM Name *", asm_options, key=f"asm_name_{fc}")

df_f3 = df_f2[df_f2['Asm Name'] == selected_asm] if selected_asm != "Select..." and not df_f2.empty else pd.DataFrame(columns=df_master.columns)

sm_options = ["Select..."] + sorted([x for x in df_f3['Sm Name'].unique() if x and x != 'nan']) if not df_f3.empty else ["Select..."]
selected_sm = st.selectbox("SM Name *", sm_options, key=f"sm_name_{fc}")

df_f4 = df_f3[df_f3['Sm Name'] == selected_sm] if selected_sm != "Select..." and not df_f3.empty else pd.DataFrame(columns=df_master.columns)

dist_options = ["Select..."] + sorted([x for x in df_f4['Distributor Name, Town DRB Code'].dropna().unique() if x and x != 'nan']) if not df_f4.empty else ["Select..."]
selected_dist = st.selectbox("Distributor Name & Code *", dist_options, key=f"dist_name_{fc}")

df_f5 = df_f4[df_f4['Distributor Name, Town DRB Code'] == selected_dist] if selected_dist != "Select..." and not df_f4.empty else pd.DataFrame(columns=df_master.columns)

spoke_options = ["Select..."] + sorted([x for x in df_f5['Spoke Name, Town Spoke Code'].dropna().unique() if x and x != 'nan']) if not df_f5.empty else ["Select..."]
selected_spoke = st.selectbox("Spoke Name & Code *", spoke_options, key=f"spoke_name_{fc}")

entered_village = st.text_input("Village Name * (Type here)", key=f"village_name_{fc}")
coverage_status = st.selectbox("Covered / Uncovered *", ["Select...", "Covered", "Uncovered"], key=f"coverage_status_{fc}")
outlet_count = st.number_input("Outlet In Village", min_value=0, value=0, step=1, key=f"outlet_count_{fc}")

st.markdown("---")
st.subheader("🌐 Location Capture & OpenStreetMap")
st.write("Click below to capture GPS location:")

loc = streamlit_geolocation()

if loc and loc.get('latitude') and loc.get('longitude'):
    lat = loc['latitude']
    lon = loc['longitude']
    acc = loc.get('accuracy', 0)
    
    st.success(f"📍 Location Captured! Accuracy: ({acc:.1f} meters)")
    
    m = folium.Map(location=[lat, lon], zoom_start=17, tiles="OpenStreetMap")
    folium.Marker(
        [lat, lon],
        popup=f"<b>{entered_village if entered_village else 'Survey Location'}</b><br>Accuracy: {acc:.1f}m",
        tooltip="Captured Location",
        icon=folium.Icon(color="blue", icon="info-sign")
    ).add_to(m)
    
    st_folium(m, width=700, height=400, key=f"map_{fc}")

if 'submitted_successfully' not in st.session_state:
    st.session_state.submitted_successfully = False

if st.session_state.submitted_successfully:
    st.success("🎉 Form Successfully Saved & Synced to Google Sheet!")
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
            
            new_record = {
                'UNIQUE_ID': current_uid,
                'Date': sub_date,
                'Time': sub_time,
                'RD Name': selected_rd,
                'STL NAMe': selected_se,
                'ASM Name': selected_asm,
                'SM Name': selected_sm,
                'Village Name': entered_village.strip(),
                'Covered/Uncovered': coverage_status,
                'Distributor Name & Code': selected_dist,
                'Spoke Name & Code': selected_spoke,
                'Outlet In Village': outlet_count,
                'Location': location_str
            }
            
            try:
                file_name_json = os.path.join(backup_dir, f"{current_uid}.json")
                with open(file_name_json, "w", encoding="utf-8") as f:
                    json.dump(new_record, f, ensure_ascii=False, indent=4)
                
                success, msg = sync_to_google_sheet(new_record)
                
                if success:
                    st.session_state.submitted_successfully = True
                    st.rerun()
                else:
                    st.error(f"❌ Google Sheet Sync Error: {msg}")
            except Exception as ex:
                st.error(f"❌ Error saving form: {ex}")
