import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from streamlit_geolocation import streamlit_geolocation
from streamlit_folium import st_folium
import folium
import os
import json

st.set_page_config(page_title="Village Coverage 2026 Form", layout="centered")

st.title("📍 Village Coverage 2026 Form")

file_path = "Village 2026.xlsx"
backup_dir = "survey_backups"

# Backup folder create karna agar na ho
if not os.path.exists(backup_dir):
    os.makedirs(backup_dir)

# 100% Safe Auto-Sync Function (Purana & Back-date data kabhi delete nahi hoga)
def auto_sync_json_to_excel():
    try:
        json_files = [os.path.join(backup_dir, f) for f in os.listdir(backup_dir) if f.endswith('.json')]
        
        # Pehle Excel ki master file aur sheet load karenge
        xls = pd.ExcelFile(file_path)
        existing_df = pd.read_excel(xls, sheet_name="Village Coverage 2026")
        
        if json_files:
            all_records = []
            for jf in json_files:
                with open(jf, 'r', encoding='utf-8') as jf_file:
                    all_records.append(json.load(jf_file))
            
            new_df = pd.DataFrame(all_records)
            
            # Purana Excel data aur saari JSON files ka data combine karke duplicates hata denge
            combined_df = pd.concat([existing_df, new_df]).drop_duplicates(subset=['UNIQUE_ID'], keep='first')
        else:
            combined_df = existing_df
            
        # File ko overwrite (mode='w') karenge taaki overlapping ki problem na aaye aur sara back-date data safe rahe
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            for sheet in xls.sheet_names:
                if sheet != "Village Coverage 2026":
                    df_sheet = pd.read_excel(xls, sheet_name=sheet)
                    df_sheet.to_excel(writer, sheet_name=sheet, index=False)
            combined_df.to_excel(writer, sheet_name="Village Coverage 2026", index=False)
            
    except Exception as e:
        print(f"Auto-sync error: {e}")

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

# Har naye form ke liye unique ID jo kabhi overlap nahi hogi
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

# 1. RD Name
rd_options = ["Select..."] + sorted([x for x in df_master['RD NAME'].unique() if x and x != 'nan'])
selected_rd = st.selectbox("RD Name *", rd_options, key=f"rd_name_{fc}")

if selected_rd != "Select...":
    df_f1 = df_master[df_master['RD NAME'] == selected_rd]
else:
    df_f1 = pd.DataFrame(columns=df_master.columns)

# 2. S.E Name
se_options = ["Select..."] + sorted([x for x in df_f1['S.E Name'].unique() if x and x != 'nan']) if not df_f1.empty else ["Select..."]
selected_se = st.selectbox("STL / S.E Name *", se_options, key=f"se_name_{fc}")

if selected_se != "Select..." and not df_f1.empty:
    df_f2 = df_f1[df_f1['S.E Name'] == selected_se]
else:
    df_f2 = pd.DataFrame(columns=df_master.columns)

# 3. ASM Name
asm_options = ["Select..."] + sorted([x for x in df_f2['Asm Name'].unique() if x and x != 'nan']) if not df_f2.empty else ["Select..."]
selected_asm = st.selectbox("ASM Name *", asm_options, key=f"asm_name_{fc}")

if selected_asm != "Select..." and not df_f2.empty:
    df_f3 = df_f2[df_f2['Asm Name'] == selected_asm]
else:
    df_f3 = pd.DataFrame(columns=df_master.columns)

# 4. SM Name
sm_options = ["Select..."] + sorted([x for x in df_f3['Sm Name'].unique() if x and x != 'nan']) if not df_f3.empty else ["Select..."]
selected_sm = st.selectbox("SM Name *", sm_options, key=f"sm_name_{fc}")

if selected_sm != "Select..." and not df_f3.empty:
    df_f4 = df_f3[df_f3['Sm Name'] == selected_sm]
else:
    df_f4 = pd.DataFrame(columns=df_master.columns)

# 5. Distributor Name & Code
dist_options = ["Select..."] + sorted([x for x in df_f4['Distributor Name, Town DRB Code'].dropna().unique() if x and x != 'nan']) if not df_f4.empty else ["Select..."]
selected_dist = st.selectbox("Distributor Name & Code *", dist_options, key=f"dist_name_{fc}")

if selected_dist != "Select..." and not df_f4.empty:
    df_f5 = df_f4[df_f4['Distributor Name, Town DRB Code'] == selected_dist]
else:
    df_f5 = pd.DataFrame(columns=df_master.columns)

# 6. Spoke Name & Code
spoke_options = ["Select..."] + sorted([x for x in df_f5['Spoke Name, Town Spoke Code'].dropna().unique() if x and x != 'nan']) if not df_f5.empty else ["Select..."]
selected_spoke = st.selectbox("Spoke Name & Code *", spoke_options, key=f"spoke_name_{fc}")

# 7. Village Name
entered_village = st.text_input("Village Name * (Type here)", key=f"village_name_{fc}")

# 8. Covered / Uncovered
coverage_status = st.selectbox("Covered / Uncovered *", ["Select...", "Covered", "Uncovered"], key=f"coverage_status_{fc}")

# 9. Outlet In Village
outlet_count = st.number_input("Outlet In Village", min_value=0, value=0, step=1, key=f"outlet_count_{fc}")

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

if 'submitted_successfully' not in st.session_state:
    st.session_state.submitted_successfully = False

if st.session_state.submitted_successfully:
    st.success("🎉 Form Successfully Saved and Auto-Synced!")
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
                # 1. Pehle safe JSON backup mein save hoga
                file_name_json = os.path.join(backup_dir, f"{current_uid}.json")
                with open(file_name_json, "w", encoding="utf-8") as f:
                    json.dump(new_record, f, ensure_ascii=False, indent=4)
                
                # 2. Phir sara purana aur naya data safely merge hokar Excel mein update ho jayega
                auto_sync_json_to_excel()
                
                st.session_state.submitted_successfully = True
                st.rerun()
            except Exception as ex:
                st.error(f"❌ Error saving form: {ex}")

# --- ADMIN PANEL ---
st.sidebar.markdown("---")
st.sidebar.subheader("🔒 Admin Download Panel")
admin_password = st.sidebar.text_input("Enter Password to Download", type="password")

if admin_password == "slmg2026":
    st.sidebar.success("✅ Access Granted (Auto-Sync Active)")
    
    try:
        # Download se pehle ensure kar lenge ki saara backup data Excel mein merged hai
        auto_sync_json_to_excel()
        
        with open(file_path, "rb") as f:
            excel_data = f.read()
        st.sidebar.download_button(
            label="📥 Download Full Survey Excel",
            data=excel_data,
            file_name="Village_Coverage_Survey_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.sidebar.info("Excel file not ready yet.")
elif admin_password != "":
    st.sidebar.error("❌ Incorrect Password")