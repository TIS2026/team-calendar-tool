import streamlit as st
import msal
import requests
import csv
import io
from datetime import datetime

import os

# Official Azure App Registration
CLIENT_ID = "afcd0889-a697-4245-9746-be99a2c64a57"
TENANT_ID = "3204476b-b2c3-4b2a-9040-c9319eafdacd"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Calendars.Read.Shared", "User.Read"]
CACHE_FILE = "token_cache.bin"

st.set_page_config(page_title="Team Shared Calendar Tool", layout="wide")
st.title("Team Shared Calendar Tool")

st.markdown("""
*Cloud-based sync directly from Microsoft Graph API. Compatible with New Outlook and Web.*
""")

def _load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache

def _save_cache(cache):
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

def get_msal_app():
    return msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=_load_cache()
    )

if 'access_token' not in st.session_state:
    st.session_state.access_token = None
    # Try to authenticate silently using the local cache
    app = get_msal_app()
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(app.token_cache)
            st.session_state.access_token = result["access_token"]

if 'device_flow' not in st.session_state:
    st.session_state.device_flow = None

def start_auth():
    app = get_msal_app()
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" in flow:
        st.session_state.device_flow = flow
        st.rerun()
    else:
        st.error(f"Failed to create device flow: {flow}")

def complete_auth():
    app = get_msal_app()
    with st.spinner("Waiting for you to complete login in your browser..."):
        result = app.acquire_token_by_device_flow(st.session_state.device_flow)
        if "access_token" in result:
            _save_cache(app.token_cache)
            st.session_state.access_token = result["access_token"]
            st.session_state.device_flow = None
            st.rerun()
        else:
            st.error(f"Authentication failed: {result.get('error_description', result.get('error'))}")
            st.session_state.device_flow = None

if not st.session_state.access_token:
    if not st.session_state.device_flow:
        st.write("Please authenticate to access your calendars.")
        if st.button("Log in with Microsoft"):
            start_auth()
    else:
        flow = st.session_state.device_flow
        st.warning("⚠️ **Action Required**")
        st.markdown(f"1. Open this link in your browser: **[{flow['verification_uri']}]({flow['verification_uri']})**")
        st.markdown(f"2. Enter the code: **`{flow['user_code']}`**")
        st.markdown("3. Complete the login process. This page will automatically update once you finish.")
        complete_auth()
    st.stop()

headers = {'Authorization': 'Bearer ' + st.session_state.access_token}

@st.cache_data(ttl=300)
def fetch_calendars():
    calendars = []
    group_url = 'https://graph.microsoft.com/v1.0/me/calendarGroups'
    
    while group_url:
        groups_resp = requests.get(group_url, headers=headers)
        if groups_resp.status_code == 200:
            groups_data = groups_resp.json()
            groups = groups_data.get('value', [])
            
            for group in groups:
                # Ignore all calendars not in the "People's Calendars" group
                if group['name'] != "People's Calendars":
                    continue
                    
                cal_url = f"https://graph.microsoft.com/v1.0/me/calendarGroups/{group['id']}/calendars?$top=100"
                while cal_url:
                    cal_resp = requests.get(cal_url, headers=headers)
                    if cal_resp.status_code == 200:
                        cal_data = cal_resp.json()
                        cals = cal_data.get('value', [])
                        
                        for c in cals:
                            owner = c.get('owner') or {}
                            if "abinash.dash" in owner.get('address', '').lower():
                                continue
                            
                            # Clean up the calendar name by removing standard Outlook group prefixes
                            cal_name = f"{group['name']} / {c['name']}"
                            cal_name = cal_name.replace("My Calendars / ", "")
                            cal_name = cal_name.replace("Shared Calendars / ", "")
                            cal_name = cal_name.replace("People's Calendars / ", "")
                            cal_name = cal_name.replace("Other Calendars / ", "")
                            
                            calendars.append({
                                "id": c['id'],
                                "name": cal_name,
                                "owner": owner.get('name', 'Unknown')
                            })
                        
                        cal_url = cal_data.get('@odata.nextLink')
                    else:
                        st.error(f"Failed to fetch calendars in group {group['name']}: {cal_resp.text}")
                        break
                        
            group_url = groups_data.get('@odata.nextLink')
        else:
            st.error(f"Failed to fetch calendar groups: {groups_resp.text}")
            break
            
    return calendars

def fetch_events(calendar_id, start_dt, end_dt):
    url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/calendarView"
    params = {
        "startDateTime": start_dt.isoformat() + "Z",
        "endDateTime": end_dt.isoformat() + "Z",
        "$top": 100,
        "$select": "subject,start,end,organizer,showAs"
    }
    
    # Request times in IST instead of UTC
    event_headers = headers.copy()
    event_headers['Prefer'] = 'outlook.timezone="India Standard Time"'
    
    events = []
    while url:
        resp = requests.get(url, headers=event_headers, params=params)
        if resp.status_code == 200:
            data = resp.json()
            for item in data.get('value', []):
                start_str = item.get('start', {}).get('dateTime', '')
                end_str = item.get('end', {}).get('dateTime', '')
                try:
                    start_val = datetime.fromisoformat(start_str.split('.')[0])
                    end_val = datetime.fromisoformat(end_str.split('.')[0])
                except ValueError:
                    continue
                events.append({
                    "Subject": item.get('subject'),
                    "Start": start_val,
                    "End": end_val,
                    "Organizer": item.get('organizer', {}).get('emailAddress', {}).get('name'),
                    "ShowAs": item.get('showAs')
                })
            url = data.get('@odata.nextLink')
            params = None
        else:
            st.error(f"Failed to fetch events: {resp.text}")
            break
    return events

st.sidebar.header("Configuration")
date_range = st.sidebar.date_input("Select Date Range", value=[datetime.now(), datetime.now()])

with st.spinner("Loading calendars from Microsoft Graph..."):
    calendars = fetch_calendars()

if not calendars:
    st.warning("No calendars found.")
    if st.button("Log out"):
        st.session_state.access_token = None
        st.rerun()
    st.stop()

if 'selected_cals' not in st.session_state:
    st.session_state.selected_cals = []

def select_all():
    cal_names = [f"{c['name']}" for c in calendars]
    st.session_state.selected_cals = cal_names

st.sidebar.button("Select All", on_click=select_all)

cal_options = {c['name']: c['id'] for c in calendars}
selected_cals = st.sidebar.multiselect(
    "Select Calendars to analyze", 
    options=list(cal_options.keys()),
    key="selected_cals"
)

if st.button("Analyze Calendars") and len(date_range) == 2 and selected_cals:
    start_date, end_date = date_range
    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.max.time())
    
    all_events = []
    with st.spinner("Fetching events..."):
        for cal_name in selected_cals:
            cal_id = cal_options[cal_name]
            events = fetch_events(cal_id, start_dt, end_dt)
            for e in events:
                e['Calendar'] = cal_name
            all_events.extend(events)
            
    if not all_events:
        st.info("No events found in the selected date range.")
        st.stop()
        
    all_events.sort(key=lambda x: x['Start'])
    
    st.subheader("All Events")
    
    display_data = [{
        "Subject": e['Subject'],
        "Start": e['Start'].strftime('%Y-%m-%d %H:%M'),
        "End": e['End'].strftime('%Y-%m-%d %H:%M'),
        "Calendar": e['Calendar'],
        "Organizer": e['Organizer'],
        "ShowAs": e['ShowAs']
    } for e in all_events]
    st.table(display_data)
    
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=["Subject", "Start", "End", "Calendar", "Organizer", "ShowAs"])
    writer.writeheader()
    writer.writerows(display_data)
    csv_bytes = output.getvalue().encode('utf-8')
    
    st.download_button(
        label="Download consolidated CSV",
        data=csv_bytes,
        file_name='consolidated_calendars.csv',
        mime='text/csv',
    )
    
    st.subheader("Scheduling Conflicts")
    conflicts = []
    
    # Group events by calendar to only check conflicts within the same calendar
    events_by_cal = {}
    for e in all_events:
        cal = e['Calendar']
        if cal not in events_by_cal:
            events_by_cal[cal] = []
        events_by_cal[cal].append(e)
        
    for cal, evs in events_by_cal.items():
        # Ensure events are sorted by start time
        evs.sort(key=lambda x: x['Start'])
        for i in range(len(evs)):
            for j in range(i+1, len(evs)):
                e1 = evs[i]
                e2 = evs[j]
                
                if e1['ShowAs'] == 'free' or e2['ShowAs'] == 'free':
                    continue
                
                # Ignore exact duplicates (same subject and same times) which are likely Graph API glitches
                if e1['Subject'] == e2['Subject'] and e1['Start'] == e2['Start'] and e1['End'] == e2['End']:
                    continue
                
                if e2['Start'] >= e1['End']:
                    break
                
                # Calculate the exact overlapping time period
                overlap_start = max(e1['Start'], e2['Start'])
                overlap_end = min(e1['End'], e2['End'])
                overlap_str = f"{overlap_start.strftime('%Y-%m-%d %H:%M')} to {overlap_end.strftime('%H:%M')}"
                
                conflicts.append({
                    "Calendar": cal,
                    "Conflict Time Period": overlap_str,
                    "Event 1": e1['Subject'],
                    "Event 1 Time": f"{e1['Start'].strftime('%H:%M')} - {e1['End'].strftime('%H:%M')}",
                    "Event 2": e2['Subject'],
                    "Event 2 Time": f"{e2['Start'].strftime('%H:%M')} - {e2['End'].strftime('%H:%M')}"
                })
            
    if conflicts:
        st.warning(f"Found {len(conflicts)} potential conflicts!")
        st.table(conflicts)
        
        conflicts_output = io.StringIO()
        conflicts_writer = csv.DictWriter(conflicts_output, fieldnames=["Calendar", "Conflict Time Period", "Event 1", "Event 1 Time", "Event 2", "Event 2 Time"])
        conflicts_writer.writeheader()
        conflicts_writer.writerows(conflicts)
        conflicts_csv_bytes = conflicts_output.getvalue().encode('utf-8')
        
        st.download_button(
            label="Download Conflicts CSV",
            data=conflicts_csv_bytes,
            file_name='scheduling_conflicts.csv',
            mime='text/csv',
        )
    else:
        st.success("No scheduling conflicts found in the selected date range!")

if st.sidebar.button("Log out"):
    st.session_state.access_token = None
    st.rerun()
