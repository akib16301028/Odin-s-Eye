import pandas as pd
import streamlit as st
from datetime import datetime
import requests  # For sending Telegram notifications

# Function to load Zone to Name mapping from USER NAME.xlsx
def load_zone_name_mapping():
    user_name_file = "USER NAME.xlsx"
    try:
        user_name_df = pd.read_excel(user_name_file)
        return dict(zip(user_name_df['Zone'], user_name_df['Name']))
    except Exception as e:
        st.error(f"Failed to load zone-name mapping from {user_name_file}. Error: {e}")
        return {}

# Function to send Telegram notification
def send_telegram_notification(message, bot_token, chat_id):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "Markdown"  # Use Markdown for plain text
    }
    response = requests.post(url, json=payload)
    return response.status_code == 200

# Streamlit app
st.title('Odin-s-Eye')

site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Alarms Excel", type=["xlsx"])

if "filter_time" not in st.session_state:
    st.session_state.filter_time = datetime.now().time()
if "filter_date" not in st.session_state:
    st.session_state.filter_date = datetime.now().date()
if "status_filter" not in st.session_state:
    st.session_state.status_filter = "All"

if site_access_file and rms_file and current_alarms_file:
    site_access_df = pd.read_excel(site_access_file)
    rms_df = pd.read_excel(rms_file, header=2)
    current_alarms_df = pd.read_excel(current_alarms_file, header=2)

    merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

    # Filter inputs (date and time)
    selected_date = st.date_input("Select Date", value=st.session_state.filter_date)
    selected_time = st.time_input("Select Time", value=st.session_state.filter_time)

    # Button to clear filters
    if st.button("Clear Filters"):
        st.session_state.filter_date = datetime.now().date()
        st.session_state.filter_time = datetime.now().time()
        st.session_state.status_filter = "All"

    # Update session state only when the user changes time or date
    if selected_date != st.session_state.filter_date:
        st.session_state.filter_date = selected_date
    if selected_time != st.session_state.filter_time:
        st.session_state.filter_time = selected_time

    # Combine selected date and time into a datetime object
    filter_datetime = datetime.combine(st.session_state.filter_date, st.session_state.filter_time)

    # Process mismatches
    mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
    mismatches_df['Start Time'] = pd.to_datetime(mismatches_df['Start Time'], errors='coerce')
    filtered_mismatches_df = mismatches_df[mismatches_df['Start Time'] > filter_datetime]

    # Process matches
    matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)

    # Apply filtering conditions
    status_filter_condition = matched_df['Status'] == st.session_state.status_filter if st.session_state.status_filter != "All" else True
    time_filter_condition = (matched_df['Start Time'] > filter_datetime) | (matched_df['End Time'] > filter_datetime)

    # Apply filters to matched data
    filtered_matched_df = matched_df[status_filter_condition & time_filter_condition]

    # Add the status filter dropdown right before the matched sites table
    status_filter = st.selectbox("Filter by Status", options=["All", "Valid", "Expired"], index=0)

    # Update session state for status filter
    if status_filter != st.session_state.status_filter:
        st.session_state.status_filter = status_filter

    # Load Zone to Name mapping
    zone_name_mapping = load_zone_name_mapping()

    # Move the "Send Telegram Notification" button to the top
    if st.button("Send Telegram Notification"):
        zones = filtered_mismatches_df['Zone'].unique()
        bot_token = "7145427044:AAGb-CcT8zF_XYkutnqqCdNLqf6qw4KgqME"
        chat_id = "-4537588687"

        for zone in zones:
            zone_df = filtered_mismatches_df[filtered_mismatches_df['Zone'] == zone]
            message = f"*Door Open Notification*\n\n*{zone}*"
            
            # If there is a corresponding name for the zone, add it to the message
            name = zone_name_mapping.get(zone, "Unknown")
            message += f"\n\n{zone} is managed by {name}:\n"

            site_aliases = zone_df['Site Alias'].unique()
            for site_alias in site_aliases:
                site_df = zone_df[zone_df['Site Alias'] == site_alias]
                message += f"#{site_alias}\n"
                for _, row in site_df.iterrows():
                    end_time_display = row['End Time'] if row['End Time'] != 'Not Closed' else 'Not Closed'
                    message += f"Start Time: {row['Start Time']} End Time: {end_time_display}\n"
                message += "\n"

            # Add the no site access request message at the end
            message += f"\n@{name}, no site access request has been found for these door open alarms. Please take care and share us update."

            if send_telegram_notification(message, bot_token, chat_id):
                st.success(f"Notification for zone '{zone}' sent successfully!")
            else:
                st.error(f"Failed to send notification for zone '{zone}'.")

    # Display mismatches
    if not filtered_mismatches_df.empty:
        st.write(f"Mismatched Sites (After {filter_datetime}) grouped by Cluster and Zone:")
        display_grouped_data(filtered_mismatches_df, "Filtered Mismatched Sites")
    else:
        st.write(f"No mismatches found after {filter_datetime}. Showing all mismatched sites.")
        display_grouped_data(mismatches_df, "All Mismatched Sites")

    # Display matched sites
    display_matched_sites(filtered_matched_df)

else:
    st.write("Please upload all required files.")
