import streamlit as st
import msal
import requests
import csv
import io
import difflib
from datetime import datetime, timedelta
import pandas as pd
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
                # Removed all group filters to guarantee no coworkers are missed
                # (You can just ignore any personal calendars you don't want to analyze)
                    
                cal_url = f"https://graph.microsoft.com/v1.0/me/calendarGroups/{group['id']}/calendars?$top=100"
                while cal_url:
                    cal_resp = requests.get(cal_url, headers=headers)
                    if cal_resp.status_code == 200:
                        cal_data = cal_resp.json()
                        cals = cal_data.get('value', [])
                        
                        for c in cals:
                            if c['name'] in ["Birthdays", "Calendar"]:
                                continue
                            
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

@st.cache_data(ttl=300, show_spinner=False)
def fetch_events(calendar_id, start_dt, end_dt, include_canceled=False):
    url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/calendarView"
    params = {
        "startDateTime": start_dt.isoformat() + "Z",
        "endDateTime": end_dt.isoformat() + "Z",
        "$top": 100,
        "$select": "subject,start,end,organizer,showAs,isCancelled"
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
                is_cancelled = item.get('isCancelled', False)
                if not is_cancelled and str(item.get('subject') or "").startswith("Canceled:"):
                    is_cancelled = True
                    
                if not include_canceled and is_cancelled:
                    continue
                    
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
                    "OrganizerEmail": item.get('organizer', {}).get('emailAddress', {}).get('address', '').lower(),
                    "ShowAs": item.get('showAs')
                })
            url = data.get('@odata.nextLink')
            params = None
        else:
            st.error(f"Failed to fetch events: {resp.text}")
            break
    return events


@st.cache_data(ttl=3600, show_spinner=False)
def load_excel_data():
    xls = pd.ExcelFile('Course wise Mentor Skillset map.xlsx')
    
    dadar_df = pd.read_excel(xls, 'Dadar')
    bandra_df = pd.read_excel(xls, 'Bandra')
    
    courses = {'Dadar': {}, 'Bandra': {}, 'Online': {}}
    
    def parse_courses(df, center):
        for _, row in df.iterrows():
            course_name = row.get('Course Name')
            if pd.isna(course_name): continue
            mentors = []
            for col in df.columns:
                if str(col).startswith('Mentor'):
                    m = row[col]
                    if pd.notna(m) and str(m).strip():
                        mentors.append(str(m).strip())
            courses[center][course_name] = mentors
            
            if course_name not in courses['Online']:
                courses['Online'][course_name] = set()
            courses['Online'][course_name].update(mentors)
            
    parse_courses(dadar_df, 'Dadar')
    parse_courses(bandra_df, 'Bandra')
    for c in courses['Online']:
        courses['Online'][c] = list(courses['Online'][c])
        
    shifts_df = pd.read_excel(xls, 'Mentor shifts')
    shifts = {}
    for _, row in shifts_df.iterrows():
        mentor = str(row.get('Mentor', '')).strip()
        if not mentor or mentor == 'nan': continue
        shifts[mentor] = {
            'Fixed Off': str(row.get('Fixed Off', '')).strip(),
            'Other Off': str(row.get('Other Off', '')).strip(),
            'Shift times': str(row.get('Shift times', '')).strip()
        }
        
    holidays_df = pd.read_excel(xls, xls.sheet_names[-1])
    holidays = []
    for _, row in holidays_df.iterrows():
        if pd.notna(row.get('Date')):
            dt = pd.to_datetime(row['Date']).date()
            holidays.append(dt)
            
    return courses, shifts, holidays

def deduplicate_events(events_list):
    events_by_cal = {}
    for e in events_list:
        cal = e['Calendar']
        if cal not in events_by_cal:
            events_by_cal[cal] = []
        events_by_cal[cal].append(e)
        
    deduped = []
    for cal, evs in events_by_cal.items():
        evs.sort(key=lambda x: x['Start'])
        valid_evs = []
        for e in evs:
            is_dup = False
            for v in valid_evs:
                if max(e['Start'], v['Start']) < min(e['End'], v['End']):
                    subj1 = (e['Subject'] or "").lower()
                    subj2 = (v['Subject'] or "").lower()
                    if subj1 == subj2:
                        is_dup = True
                        break
                    if difflib.SequenceMatcher(None, subj1, subj2).ratio() > 0.7 and e['Start'] == v['Start']:
                        is_dup = True
                        break
            if not is_dup:
                valid_evs.append(e)
        deduped.extend(valid_evs)
    return deduped


st.sidebar.header("Navigation")
nav_mode = st.sidebar.radio("Go to", ["Smart Scheduler", "Raw Events", "Scheduling Conflicts"])

with st.spinner("Loading calendars from Microsoft Graph..."):
    calendars = fetch_calendars()

if not calendars:
    st.warning("No calendars found.")
    if st.button("Log out"):
        st.session_state.access_token = None
        st.rerun()
    st.stop()

cal_options = {c['name']: c['id'] for c in calendars}

if 'selected_cals' not in st.session_state:
    st.session_state.selected_cals = []

def select_all():
    cal_names = [f"{c['name']}" for c in calendars]
    st.session_state.selected_cals = cal_names

if nav_mode in ["Raw Events", "Scheduling Conflicts"]:
    st.sidebar.header("Configuration")
    date_range = st.sidebar.date_input("Select Date Range", value=[datetime.now(), datetime.now()])
    include_canceled = st.sidebar.checkbox("Include Canceled Meetings", value=False)
    st.sidebar.button("Select All", on_click=select_all)
    selected_cals = st.sidebar.multiselect(
        "Select Calendars to analyze", 
        options=list(cal_options.keys()),
        key="selected_cals"
    )
    if len(date_range) == 2:
        start_date, end_date = date_range
        start_dt = datetime.combine(start_date, datetime.min.time())
        end_dt = datetime.combine(end_date, datetime.max.time())
    else:
        start_dt, end_dt = None, None
else:
    selected_cals = []
    start_dt, end_dt = None, None

try:
    courses_data, mentor_shifts, holiday_list = load_excel_data()
    all_course_names = set()
    for center_key in courses_data:
        all_course_names.update(courses_data[center_key].keys())
    all_course_names = sorted(list(all_course_names))
except Exception as e:
    st.error(f"Failed to load Excel data: {e}")
    all_course_names = []
    courses_data = {'Dadar': {}, 'Bandra': {}, 'Online': {}}
    mentor_shifts = {}
    holiday_list = []

if nav_mode == "Smart Scheduler":
    st.subheader("Smart Scheduling Engine")
    
    col1, col2 = st.columns(2)
    with col1:
        selected_course = st.selectbox("Course Name (Required)", options=[""] + all_course_names)
        center = st.radio("Center", options=["Bandra", "Dadar", "Online"], horizontal=True)
        total_hours = st.number_input("Number of hours to schedule (Required)", min_value=0.5, step=0.5, value=10.0)
        
        default_start = datetime.now().date() + timedelta(days=1)
        sched_start_date = st.date_input("Start Date", value=default_start)
        sched_end_date = st.date_input("End Date (Optional, acts as a hard deadline)", value=None)
            
    with col2:
        st.markdown("**Session Constraints**")
        bypass_limit = st.checkbox("Bypass 2-hour daily limit")
        
        import numpy as np
        if bypass_limit:
            max_duration = st.number_input("Maximum number of hours in a session", min_value=1.0, value=3.0, step=0.5)
            duration_options = list(np.arange(1.0, max_duration + 0.5, 0.5))
        else:
            duration_options = [1.0, 1.5, 2.0]
            
        pref_duration = st.selectbox("Preferred Session Duration (Hours)", options=duration_options, index=duration_options.index(2.0) if 2.0 in duration_options else 0)
        
        st.markdown("**Session Times**")
        p_t_col1, p_t_col2 = st.columns(2)
        with p_t_col1:
            pref_start_time = st.time_input("Preferred Start Bound", value=None)
        with p_t_col2:
            pref_end_time = st.time_input("Preferred End Bound", value=None)
            
        st.markdown("**Weekdays**")
        weekdays = st.multiselect("Preferred Weekdays (Optional)", 
            options=["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
        
    with st.expander("Advanced Fallback Options (Optional)"):
        st.markdown("Use these fields to allow the algorithm to gracefully fall back to alternative options if your preferred setup isn't possible.")
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            filler_weekdays = st.multiselect("Fallback Weekdays", 
                options=["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
        with f_col2:
            f_t_col1, f_t_col2 = st.columns(2)
            with f_t_col1:
                filler_start_time = st.time_input("Fallback Start Bound", value=None)
            with f_t_col2:
                filler_end_time = st.time_input("Fallback End Bound", value=None)
        with f_col3:
            filler_options = [d for d in duration_options if d != pref_duration]
            filler_durations = st.multiselect("Fallback Durations", options=filler_options)
            
    if st.button("Find Available Schedules", type="primary"):
        if not selected_course:
            st.error("Please select a Course Name.")
        elif total_hours <= 0:
            st.error("Number of hours to schedule must be greater than 0.")
        else:
            with st.spinner("Analyzing mentor schedules across multi-dimensional profiles..."):
                mentors_needed = courses_data.get(center, {}).get(selected_course, [])
                if not mentors_needed:
                    st.warning(f"No mentors found for {selected_course} at {center}.")
                else:
                    available_cals = []
                    for m in mentors_needed:
                        matched = False
                        
                        import difflib
                        import re
                        m_parts = [p for p in re.split(r'[^a-zA-Z0-9]', m.lower()) if p]
                        
                        for cal_name in cal_options:
                            c_clean = cal_name.lower()
                            c_parts = [p for p in re.split(r'[^a-zA-Z0-9]', c_clean) if p]
                            
                            is_fuzzy_match = True
                            if not m_parts or not c_parts:
                                is_fuzzy_match = False
                            else:
                                for mp in m_parts:
                                    best_ratio = max([difflib.SequenceMatcher(None, mp, cp).ratio() for cp in c_parts] + [0])
                                    if best_ratio < 0.80:
                                        is_fuzzy_match = False
                                        break
                                        
                            if (m.lower() in c_clean or 
                                m.lower().replace(" ", "") in c_clean.replace(" ", "") or 
                                is_fuzzy_match):
                                available_cals.append((m, cal_name, cal_options[cal_name]))
                                matched = True
                                break
                        if not matched:
                            st.warning(f"Could not find a connected calendar for mentor: **{m}**.\n\n"
                                       f"*Troubleshooting:*\n"
                                       f"1. Open your Outlook Calendar.\n"
                                       f"2. Ensure you have explicitly added '{m}' to your 'Shared Calendars' or 'People's Calendars' list on the left sidebar.\n"
                                       f"3. Ensure the name spelled in Outlook matches the Excel sheet.")
                    
                    if available_cals:
                        import math
                        import re
                        from datetime import time
                        from collections import deque
                        
                        def get_optimal_schedule_mix(total_h, pref_d, fillers):
                            max_pref = int(total_h // pref_d)
                            for p_count in range(max_pref, -1, -1):
                                rem = total_h - (p_count * pref_d)
                                if rem == 0:
                                    return [pref_d] * p_count
                                if not fillers: continue
                                q = deque([ (rem, []) ])
                                valid_fillers = []
                                while q:
                                    curr_rem, path = q.popleft()
                                    if curr_rem == 0:
                                        valid_fillers = path
                                        break
                                    if curr_rem < 0:
                                        continue
                                    for f in fillers:
                                        if not path or f <= path[-1]:
                                            q.append((curr_rem - f, path + [f]))
                                if valid_fillers:
                                    return [pref_d] * p_count + valid_fillers
                            return None

                        pure_schedule_mix = get_optimal_schedule_mix(total_hours, pref_duration, [])
                        mixed_schedule_mix = get_optimal_schedule_mix(total_hours, pref_duration, filler_durations) if filler_durations else None
                        
                        f_s_time = filler_start_time if filler_start_time else pref_start_time
                        f_e_time = filler_end_time if filler_end_time else pref_end_time
                        f_w_days = (weekdays or []) + (filler_weekdays or []) if filler_weekdays else weekdays
                        
                        profiles = [
                            {"name": "Profile 1 (Pure Duration, Pref Weekdays, Pref Time)", "mix": pure_schedule_mix, "w_days": weekdays, "s_time": pref_start_time, "e_time": pref_end_time},
                            {"name": "Profile 2 (Pure Duration, Pref Weekdays, Fallback Time)", "mix": pure_schedule_mix, "w_days": weekdays, "s_time": f_s_time, "e_time": f_e_time},
                            {"name": "Profile 3 (Pure Duration, Fallback Weekdays, Pref Time)", "mix": pure_schedule_mix, "w_days": f_w_days, "s_time": pref_start_time, "e_time": pref_end_time},
                            {"name": "Profile 4 (Pure Duration, Fallback Weekdays, Fallback Time)", "mix": pure_schedule_mix, "w_days": f_w_days, "s_time": f_s_time, "e_time": f_e_time},
                        ]
                        
                        if mixed_schedule_mix and mixed_schedule_mix != pure_schedule_mix:
                            profiles += [
                                {"name": "Profile 5 (Mixed Duration, Pref Weekdays, Pref Time)", "mix": mixed_schedule_mix, "w_days": weekdays, "s_time": pref_start_time, "e_time": pref_end_time},
                                {"name": "Profile 6 (Mixed Duration, Pref Weekdays, Fallback Time)", "mix": mixed_schedule_mix, "w_days": weekdays, "s_time": f_s_time, "e_time": f_e_time},
                                {"name": "Profile 7 (Mixed Duration, Fallback Weekdays, Pref Time)", "mix": mixed_schedule_mix, "w_days": f_w_days, "s_time": pref_start_time, "e_time": pref_end_time},
                                {"name": "Profile 8 (Mixed Duration, Fallback Weekdays, Fallback Time)", "mix": mixed_schedule_mix, "w_days": f_w_days, "s_time": f_s_time, "e_time": f_e_time},
                            ]
                            
                        successful_profile = None
                        successful_schedules = []
                        target_dates_for_multi = []
                        mix_for_multi = []
                        s_time_multi, e_time_multi = None, None
                        
                        # Just need to load all events once for a very wide window to be safe
                        s_dt_global = datetime.combine(sched_start_date, datetime.min.time())
                        e_dt_global = datetime.combine(sched_end_date or (sched_start_date + timedelta(days=90)), datetime.max.time())
                        
                        all_mentor_events = {}
                        for m_name, c_name, c_id in available_cals:
                            evs = fetch_events(c_id, s_dt_global, e_dt_global, include_canceled=False)
                            busy_evs = [e for e in evs if e['ShowAs'] != 'free' and 'lunch' not in (e['Subject'] or '').lower()]
                            all_mentor_events[m_name] = busy_evs
                        
                        for p in profiles:
                            if not p["mix"]: continue
                            
                            sessions_needed = len(p["mix"])
                            current_date = sched_start_date
                            target_dates = []
                            
                            allowed_weekdays = p["w_days"]
                            
                            while len(target_dates) < sessions_needed:
                                if sched_end_date and current_date > sched_end_date: break
                                d_dt = datetime.combine(current_date, datetime.min.time())
                                day_name = d_dt.strftime('%A')
                                if allowed_weekdays and day_name not in allowed_weekdays:
                                    current_date += timedelta(days=1)
                                    continue
                                if current_date in holiday_list:
                                    current_date += timedelta(days=1)
                                    continue
                                target_dates.append(current_date)
                                current_date += timedelta(days=1)
                                
                            if len(target_dates) < sessions_needed: continue
                            
                            target_dates_for_multi = target_dates
                            mix_for_multi = p["mix"]
                            s_time_multi, e_time_multi = p["s_time"], p["e_time"]
                            
                            valid_schedules = []
                            for m_name, c_name, c_id in available_cals:
                                m_shift = mentor_shifts.get(m_name, {})
                                fixed_off = m_shift.get('Fixed Off')
                                other_off = m_shift.get('Other Off')
                                shift_times = m_shift.get('Shift times')
                                m_evs = all_mentor_events.get(m_name, [])
                                
                                mentor_valid_slots = []
                                potential_slots = []
                                for h in range(8, 21):
                                    potential_slots.append(time(h, 0))
                                    potential_slots.append(time(h, 30))
                                    
                                for p_slot in potential_slots:
                                    
                                    if p["s_time"] and p_slot < p["s_time"]: continue
                                    
                                    dynamic_target_dates = []
                                    current_date = sched_start_date
                                    
                                    # Hard limit to prevent infinite loops (e.g. searching 5 years into future)
                                    search_limit = 100 
                                    searched_days = 0
                                    
                                    while len(dynamic_target_dates) < sessions_needed and searched_days < search_limit:
                                        searched_days += 1
                                        if sched_end_date and current_date > sched_end_date: break
                                        
                                        day_name = current_date.strftime('%A')
                                        if allowed_weekdays and day_name not in allowed_weekdays:
                                            current_date += timedelta(days=1)
                                            continue
                                        if current_date in holiday_list:
                                            current_date += timedelta(days=1)
                                            continue
                                            
                                        idx = len(dynamic_target_dates)
                                        assigned_dur = p["mix"][idx]
                                        p_slot_end_dt = datetime.combine(current_date, p_slot) + timedelta(hours=assigned_dur)
                                        p_slot_end = p_slot_end_dt.time()
                                        
                                        if p_slot_end < p_slot: 
                                            current_date += timedelta(days=1); continue
                                        if p["e_time"] and p_slot_end > p["e_time"]:
                                            current_date += timedelta(days=1); continue
                                            
                                        is_weekend = day_name in ['Saturday', 'Sunday']
                                        
                                        is_off = False
                                        if fixed_off and day_name.lower() == str(fixed_off).lower(): is_off = True
                                        elif other_off and day_name.lower() in str(other_off).lower():
                                            if '2nd' in str(other_off).lower() and '4th' in str(other_off).lower():
                                                nth_week = (current_date.day - 1) // 7 + 1
                                                if nth_week in [2, 4]: is_off = True
                                        if is_off:
                                            current_date += timedelta(days=1); continue
                                            
                                        s_str = str(shift_times).lower()
                                        parts = s_str.split(',')
                                        target_part = ""
                                        if is_weekend:
                                            for part in parts:
                                                if 'weekend' in part: target_part = part
                                        else:
                                            for part in parts:
                                                if 'weekday' in part: target_part = part
                                        if not target_part: target_part = parts[0]
                                        
                                        import re
                                        t_matches = re.findall(r'(\d{1,2}(?::\d{2})?\s*(?:am|pm))', target_part.replace('-', ' to '))
                                        parsed_times = []
                                        for t_str in t_matches:
                                            t_str = t_str.replace(' ', '')
                                            try:
                                                if ':' in t_str: parsed_times.append(datetime.strptime(t_str, '%I:%M%p').time())
                                                else: parsed_times.append(datetime.strptime(t_str, '%I%p').time())
                                            except: pass
                                        if len(parsed_times) < 2:
                                            current_date += timedelta(days=1); continue
                                            
                                        m_shift_start = parsed_times[0]
                                        m_shift_end = parsed_times[-1]
                                        
                                        if p_slot < m_shift_start or p_slot_end > m_shift_end:
                                            current_date += timedelta(days=1); continue
                                            
                                        slot_start_dt = datetime.combine(current_date, p_slot)
                                        slot_end_dt = datetime.combine(current_date, p_slot_end)
                                        
                                        day_busy = [e for e in m_evs if e['Start'].date() <= current_date and e['End'].date() >= current_date]
                                        conflict = False
                                        for ev in day_busy:
                                            ev_s = max(ev['Start'], datetime.combine(current_date, time.min))
                                            ev_e = min(ev['End'], datetime.combine(current_date, time.max))
                                            if ev_s < slot_end_dt and ev_e > slot_start_dt:
                                                conflict = True
                                                break
                                                
                                        if conflict:
                                            current_date += timedelta(days=1); continue
                                            
                                        dynamic_target_dates.append(current_date)
                                        current_date += timedelta(days=1)
                                        
                                    if len(dynamic_target_dates) == sessions_needed:
                                        mentor_valid_slots.append((p_slot.strftime('%I:%M %p'), dynamic_target_dates))
                                        
                                if mentor_valid_slots:
                                    chosen_start_str, chosen_dates = mentor_valid_slots[0]
                                    chosen_start_time = datetime.strptime(chosen_start_str, '%I:%M %p').time()
                                    
                                    schedule_details = []
                                    for idx, td in enumerate(chosen_dates):
                                        assigned_dur = p["mix"][idx]
                                        end_dt = datetime.combine(datetime.today(), chosen_start_time) + timedelta(hours=assigned_dur)
                                        end_t = end_dt.time()
                                        
                                        schedule_details.append({
                                            "Session": f"Session {idx+1}",
                                            "Date": td.strftime('%Y-%m-%d (%a)'),
                                            "Time": f"{chosen_start_time.strftime('%I:%M %p')} - {end_t.strftime('%I:%M %p')}",
                                            "Duration": f"{assigned_dur} hrs",
                                            "Mentor": m_name
                                        })
                                    
                                    valid_schedules.append({
                                        "Mentor": m_name,
                                        "Schedule": schedule_details
                                    })
                                    
                            if valid_schedules:
                                successful_profile = p
                                successful_schedules = valid_schedules
                                break # Stop searching profiles! We found the best one!
                                
                        if successful_profile:
                            st.session_state['search_result'] = {
                                'type': 'single',
                                'profile': successful_profile,
                                'schedules': successful_schedules,
                                'target_dates': target_dates_for_multi
                            }
                        else:
                            multi_schedule = []
                            if target_dates_for_multi:
                                for idx, td in enumerate(target_dates_for_multi):
                                    assigned_dur = mix_for_multi[idx]
                                    day_name = td.strftime('%A')
                                    is_weekend = day_name in ['Saturday', 'Sunday']
                                    
                                    first_valid_mentor = None
                                    first_valid_start = None
                                    first_valid_end = None
                                    
                                    for m_name, c_name, c_id in available_cals:
                                        if first_valid_mentor: break
                                        
                                        m_shift = mentor_shifts.get(m_name, {})
                                        fixed_off = m_shift.get('Fixed Off')
                                        other_off = m_shift.get('Other Off')
                                        shift_times = m_shift.get('Shift times')
                                        
                                        is_off = False
                                        if fixed_off and day_name.lower() == str(fixed_off).lower(): is_off = True
                                        elif other_off and day_name.lower() in str(other_off).lower():
                                            if '2nd' in str(other_off).lower() and '4th' in str(other_off).lower():
                                                nth_week = (td.day - 1) // 7 + 1
                                                if nth_week in [2, 4]: is_off = True
                                        if is_off: continue
                                        
                                        s_str = str(shift_times).lower()
                                        parts = s_str.split(',')
                                        target_part = ""
                                        if is_weekend:
                                            for p in parts:
                                                if 'weekend' in p: target_part = p
                                        else:
                                            for p in parts:
                                                if 'weekday' in p: target_part = p
                                        if not target_part: target_part = parts[0]
                                        
                                        import re
                                        t_matches = re.findall(r'(\d{1,2}(?::\d{2})?\s*(?:am|pm))', target_part.replace('-', ' to '))
                                        parsed_times = []
                                        for t_str in t_matches:
                                            t_str = t_str.replace(' ', '')
                                            try:
                                                if ':' in t_str: parsed_times.append(datetime.strptime(t_str, '%I:%M%p').time())
                                                else: parsed_times.append(datetime.strptime(t_str, '%I%p').time())
                                            except: pass
                                        if len(parsed_times) < 2: continue
                                            
                                        m_shift_start = parsed_times[0]
                                        m_shift_end = parsed_times[-1]
                                        
                                        m_evs = all_mentor_events.get(m_name, [])
                                        day_busy = [e for e in m_evs if e['Start'].date() <= td and e['End'].date() >= td]
                                        
                                        potential_slots = []
                                        for h in range(8, 21):
                                            potential_slots.append(time(h, 0))
                                            potential_slots.append(time(h, 30))
                                            
                                        for p_slot in potential_slots:
                                            p_slot_end_dt = datetime.combine(datetime.today(), p_slot) + timedelta(hours=assigned_dur)
                                            p_slot_end = p_slot_end_dt.time()
                                            if p_slot_end < p_slot: continue
                                            if s_time_multi and p_slot < s_time_multi: continue
                                            if e_time_multi and p_slot_end > e_time_multi: continue
                                            if p_slot < m_shift_start or p_slot_end > m_shift_end: continue
                                            
                                            slot_start_dt = datetime.combine(td, p_slot)
                                            slot_end_dt = datetime.combine(td, p_slot_end)
                                            
                                            conflict = False
                                            for ev in day_busy:
                                                ev_s = max(ev['Start'], datetime.combine(td, time.min))
                                                ev_e = min(ev['End'], datetime.combine(td, time.max))
                                                if ev_s < slot_end_dt and ev_e > slot_start_dt:
                                                    conflict = True
                                                    break
                                            if not conflict:
                                                first_valid_mentor = m_name
                                                first_valid_start = p_slot
                                                first_valid_end = p_slot_end
                                                break
                                    
                                    if first_valid_mentor:
                                        multi_schedule.append({
                                            "Session": f"Session {idx+1}",
                                            "Date": td.strftime('%Y-%m-%d (%a)'),
                                            "Time": f"{first_valid_start.strftime('%I:%M %p')} - {first_valid_end.strftime('%I:%M %p')}",
                                            "Duration": f"{assigned_dur} hrs",
                                            "Mentor": first_valid_mentor
                                        })
                                    else:
                                        multi_schedule.append({
                                            "Session": f"Session {idx+1}",
                                            "Date": td.strftime('%Y-%m-%d (%a)'),
                                            "Time": "NO SLOTS AVAILABLE",
                                            "Duration": f"{assigned_dur} hrs",
                                            "Mentor": "UNSTAFFED"
                                        })
                            
                            if not target_dates_for_multi:
                                st.session_state['search_result'] = {
                                    'type': 'error',
                                    'msg': "Math Error: Could not find enough valid target dates before the deadline to accommodate the hours."
                                }
                            else:
                                st.session_state['search_result'] = {
                                    'type': 'multi',
                                    'multi_schedule': multi_schedule
                                }
                                
    # Outside the button logic, render the results from session state
    if 'search_result' in st.session_state:
        res = st.session_state['search_result']
        st.divider()
        if res['type'] == 'single':
            st.success(f"Schedule automatically generated! (Matched using {res['profile']['name']})")
            
            # Interactive Mentor Selection
            mentor_names = [s['Mentor'] for s in res['schedules']]
            
            if len(mentor_names) > 1:
                st.info(f"Great news! We found {len(mentor_names)} mentors who can accommodate this schedule perfectly.")
                selected_mentor_name = st.selectbox("Select Mentor to Assign", options=mentor_names)
            else:
                selected_mentor_name = mentor_names[0]
                st.info(f"Automatically selected Mentor: **{selected_mentor_name}** (Only available mentor for this profile)")
                
            # Render the selected schedule
            for s in res['schedules']:
                if s['Mentor'] == selected_mentor_name:
                    st.dataframe(s['Schedule'], use_container_width=True)
                    break
                    
        elif res['type'] == 'error':
            st.error(res['msg'])
            
        elif res['type'] == 'multi':
            st.warning("No single mentor is consistently available across any of the fallback profiles.")
            if st.button("Generate Multi-Mentor Schedule"):
                st.info("Generating multi-mentor schedule using the most flexible constraints (Profile 8)...")
                st.dataframe(res['multi_schedule'], use_container_width=True)

elif nav_mode == "Raw Events":

    if not selected_cals:
        st.info("Please select calendars from the sidebar configuration to view raw events.")
    elif st.button("Fetch All Events"):
        all_events = []
        with st.spinner("Fetching events..."):
            for cal_name in selected_cals:
                cal_id = cal_options[cal_name]
                events = fetch_events(cal_id, start_dt, end_dt, include_canceled)
                for e in events:
                    e['Calendar'] = cal_name
                all_events.extend(events)
                
        all_events = deduplicate_events(all_events)
        
        if not all_events:
            st.info("No events found in the selected date range.")
        else:
            all_events.sort(key=lambda x: x['Start'])
            
            display_data = [{
                "Date": e['Start'].strftime('%Y-%m-%d'),
                "Start_Time": e['Start'].strftime('%H:%M'),
                "End_Time": e['End'].strftime('%H:%M'),
                "Instructor": e['Calendar'],
                "Session_Meeting_Topic": e['Subject']
            } for e in all_events]
            
            output = io.StringIO()
            writer = csv.DictWriter(output, fieldnames=["Date", "Start_Time", "End_Time", "Instructor", "Session_Meeting_Topic"])
            writer.writeheader()
            writer.writerows(display_data)
            csv_bytes = output.getvalue().encode('utf-8')
            
            st.download_button(
                label="Download consolidated CSV",
                data=csv_bytes,
                file_name='consolidated_calendars.csv',
                mime='text/csv',
            )
            
            st.table(display_data)

elif nav_mode == "Scheduling Conflicts":
    if not selected_cals:
        st.info("Please select calendars from the sidebar configuration to analyze conflicts.")
    elif st.button("Analyze Conflicts"):
        all_events = []
        with st.spinner("Analyzing conflicts..."):
            for cal_name in selected_cals:
                cal_id = cal_options[cal_name]
                events = fetch_events(cal_id, start_dt, end_dt, include_canceled)
                for e in events:
                    e['Calendar'] = cal_name
                all_events.extend(events)
        
        all_events = deduplicate_events(all_events)
        
        if not all_events:
            st.info("No events found in the selected date range to check for conflicts.")
        else:
            conflicts = []
            events_by_cal = {}
            for e in all_events:
                cal = e['Calendar']
                if cal not in events_by_cal:
                    events_by_cal[cal] = []
                events_by_cal[cal].append(e)
                
            for cal, evs in events_by_cal.items():
                evs.sort(key=lambda x: x['Start'])
                for i in range(len(evs)):
                    for j in range(i+1, len(evs)):
                        e1 = evs[i]
                        e2 = evs[j]
                        
                        if e1['ShowAs'] == 'free' or e2['ShowAs'] == 'free':
                            continue
                            
                        subj1 = (e1['Subject'] or "").lower()
                        subj2 = (e2['Subject'] or "").lower()
                        if "lunch" in subj1 or "lunch" in subj2:
                            continue
                        
                        if e2['Start'] >= e1['End']:
                            break
                        
                        overlap_start = max(e1['Start'], e2['Start'])
                        overlap_end = min(e1['End'], e2['End'])
                        overlap_str = f"{overlap_start.strftime('%Y-%m-%d %H:%M')} to {overlap_end.strftime('%H:%M')}"
                        
                        org_email1 = e1.get('OrganizerEmail', '')
                        org_email2 = e2.get('OrganizerEmail', '')
                        officead_email = 'officead@theinnovationstory.com'
                        
                        is_e1_officead = (org_email1 == officead_email)
                        is_e2_officead = (org_email2 == officead_email)
                        
                        e1_subj = f"{e1['Subject']} [Blocked by Office Admin]" if is_e1_officead else e1['Subject']
                        e2_subj = f"{e2['Subject']} [Blocked by Office Admin]" if is_e2_officead else e2['Subject']
                        
                        if is_e1_officead and is_e2_officead:
                            conflict_type = "Double Blocked by Office Admin"
                        else:
                            conflict_type = "Other Conflict"
                        
                        conflicts.append({
                            "Calendar": cal,
                            "Conflict Type": conflict_type,
                            "Conflict Time Period": overlap_str,
                            "Event 1": e1_subj,
                            "Event 1 Time": f"{e1['Start'].strftime('%H:%M')} - {e1['End'].strftime('%H:%M')}",
                            "Event 2": e2_subj,
                            "Event 2 Time": f"{e2['Start'].strftime('%H:%M')} - {e2['End'].strftime('%H:%M')}"
                        })
                    
            if conflicts:
                st.warning(f"Found {len(conflicts)} potential conflicts!")
                
                conflicts.sort(key=lambda x: 0 if "Double Blocked" in x["Conflict Type"] else 1)
                
                conflicts_output = io.StringIO()
                conflicts_writer = csv.DictWriter(conflicts_output, fieldnames=["Calendar", "Conflict Type", "Conflict Time Period", "Event 1", "Event 1 Time", "Event 2", "Event 2 Time"])
                conflicts_writer.writeheader()
                conflicts_writer.writerows(conflicts)
                conflicts_csv_bytes = conflicts_output.getvalue().encode('utf-8')
                
                st.download_button(
                    label="Download Conflicts CSV",
                    data=conflicts_csv_bytes,
                    file_name='scheduling_conflicts.csv',
                    mime='text/csv',
                )
                
                st.dataframe(conflicts, use_container_width=True)
            else:
                st.success("No scheduling conflicts found in the selected date range!")

if st.sidebar.button("Log out"):
    st.session_state.access_token = None
    st.rerun()
