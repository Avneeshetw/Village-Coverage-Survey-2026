import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from streamlit_geolocation import streamlit_geolocation
from streamlit_folium import st_folium
import folium

st.set_page_config(page_title="Village Coverage 2026 Form", layout="centered")

st.title("📍 Village Coverage 2026 Form")

file_path = "Village 2026.xlsx"

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

st.subheader("Survey Details")

# IST (Indian Standard Time) correction (+5 hours 30 minutes from UTC)
ist_time = datetime.utcnow() + timedelta(hours=5, minutes=30)

col1, col2 = st.columns(2)
with col1:
    current_date = ist_time.strftime("%Y-%m-%d")
    st.text_input("Date", value=current_date, disabled=True, key="display_date")
with col2:
    current_time = ist_time.strftime("%H:%M:%S")
    st.text_input("Time", value=current_time, disabled=True, key="display_time")

# 1. RD Name
rd_options = ["Select..."] + sorted([x for x in df_master['RD NAME'].unique() if x and x != 'nan'])
selected_rd = st.selectbox("RD Name *", rd_options, key="rd_name")

if selected_rd != "Select...":
    df_f1 = df_master[df_master['RD NAME'] == selected_rd]
else:
    df_f1 = pd.DataFrame(columns=df_master.columns)

# 2. S.E Name
se_options = ["Select..."] + sorted([x for x in df_f1['S.E Name'].unique() if x and x != 'nan']) if not df_f1.empty else ["Select..."]
selected_se = st.selectbox("STL / S.E Name *", se_options, key="se_name")

if selected_se != "Select..." and not df_f1.empty:
    df_f2 = df_f1[df_f1['S.E Name'] == selected_se]
else:
    df_f2 = pd.DataFrame(columns=df_master.columns)

# 3. ASM Name
asm_options = ["Select..."] + sorted([x for x in df_f2['Asm Name'].unique() if x and x != 'nan']) if not df_f2.empty else ["Select..."]
selected_asm = st.selectbox("ASM Name *", asm_options, key="asm_name")

if selected_asm != "Select..." and not df_f2.empty:
    df_f3 = df_f2[df_f2['Asm Name'] == selected_asm]
else:
    df_f3 = pd.DataFrame(columns=df_master.columns)

# 4. SM Name
sm_options = ["Select..."] + sorted([x for x in df_f3['Sm Name'].unique() if x and x != 'nan']) if not df_f3.empty else ["Select..."]
selected_sm = st.selectbox("SM Name *", sm_options, key="sm_name")

if selected_sm != "Select..." and not df_f3.empty:
    df_f4 = df_f3[df_f3['Sm Name'] == selected_sm]
else:
    df_f4 = pd.DataFrame(columns=df_master.columns)

# 5. Distributor Name & Code
dist_options = ["Select..."] + sorted([x for x in df_f4['Distributor Name, Town DRB Code'].dropna().unique() if x and x != 'nan']) if not df_f4.empty else ["Select..."]
selected_dist = st.selectbox("Distributor Name & Code *", dist_options, key="dist_name")

if selected_dist != "Select..." and not df_f4.empty:
    df_f5 = df_f4[df_f4['Distributor Name, Town DRB Code'] == selected_dist]
else:
    df_f5 = pd.DataFrame(columns=df_master.columns)

# 6. Spoke Name & Code
spoke_options = ["Select..."] + sorted([x for x in df_f5['Spoke Name, Town Spoke Code'].dropna().unique() if x and x != 'nan']) if not df_f5.empty else ["Select..."]
selected_spoke = st.selectbox("Spoke Name & Code *", spoke_options, key="spoke_name")

# 7. Village Name (Manual Text Input)
entered_village = st.text_input("Village Name * (Type here)", key="village_name")

# 8. Covered / Uncovered
coverage_status = st.selectbox("Covered / Uncovered *", ["Select...", "Covered", "Uncovered"], key="coverage_status")

# 9. Outlet In Village
outlet_count = st.number_input("Outlet In Village", min_value=0, value=0, step=1, key="outlet_count")

st.markdown("---")
st.subheader("🌐 Location Capture & OpenStreetMap")
st.write("Click below to capture GPS location:")

loc = streamlit_geolocation()

if loc and loc.get('latitude') and loc.get('longitude'):
    lat = loc['latitude']
    lon = loc['longitude']
    acc = loc.get('accuracy', 0)
    
    st.success(f"📍 Location Captured Successfully! (Accuracy: {acc:.1f} meters)")
    
    m = folium.Map(location=[lat, lon], zoom_start=17, tiles="OpenStreetMap")
    folium.Marker(
        [lat, lon],
        popup=f"<b>{entered_village if entered_village else 'Survey Location'}</b><br>Accuracy: {acc:.1f}m",
        tooltip="Captured Location",
        icon=folium.Icon(color="blue", icon="info-sign")
    ).add_to(m)
    
    st_folium(m, width=700, height=400)

if 'submitted_successfully' not in st.session_state:
    st.session_state.submitted_successfully = False

if st.session_state.submitted_successfully:
    st.success("🎉 Form Successfully Saved and Logged to Excel!")
    if st.button("➕ Fill Next Form", type="primary"):
        st.session_state.submitted_successfully = False
        st.rerun()
else:
    if st.button("Save Form", type="primary"):
        if selected_rd == "Select..." or selected_se == "Select..." or selected_dist == "Select..." or selected_spoke == "Select..." or not entered_village.strip() or coverage_status == "Select...":
            st.error("❌ Kripya sabhi zaroori fields (* marked) bharein!")
        elif not loc or not loc.get('latitude'):
            st.error("❌ Location capture nahi hui! Kripya GPS allow karein.")
        else:
            lat = loc.get('latitude')
            lon = loc.get('longitude')
            acc = loc.get('accuracy', 0)
            location_str = f"Lat: {lat}, Lng: {lon} (Acc: {acc:.1f}m)"
            
            sub_date = ist_time.strftime("%Y-%m-%d")
            sub_time = ist_time.strftime("%H:%M:%S")
            
            new_record = {
                'UNIQUE_ID': f"UID_{int(datetime.now().timestamp())}",
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
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    existing_df = pd.read_excel(file_path, sheet_name="Village Coverage 2026")
                    updated_df = pd.concat([existing_df, pd.DataFrame([new_record])], ignore_index=True)
                    updated_df.to_excel(writer, sheet_name="Village Coverage 2026", index=False)
                
                st.session_state.submitted_successfully = True
                st.rerun()
            except Exception as ex:
                st.error(f"❌ Error saving to Excel: {ex}")