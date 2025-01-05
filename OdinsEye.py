import pandas as pd
import streamlit as st
from datetime import datetime
import requests  # For sending Telegram notifications

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to merge RMS and Current Alarms data
def merge_rms_alarms(rms_df, alarms_df):
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT  # No End Time in Current Alarms, set to NaT

    rms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
    alarms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']

    merged_df = pd.concat([rms_df[rms_columns], alarms_df[alarms_columns]], ignore_index=True)
    return merged_df

# Function to find mismatches between Site Access and merged RMS/Alarms dataset
def find_mismatches(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_comparison_df = pd.merge(merged_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('Not Closed')  # Replace NaT with Not Closed
    return mismatches_df

# Function to find matched sites and their status
def find_matched_sites(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    matched_df = pd.merge(site_access_df, merged_df, left_on='SiteName_Extracted', right_on='Site', how='inner')
    matched_df['StartDate'] = pd.to_datetime(matched_df['StartDate'], errors='coerce')
    matched_df['EndDate'] = pd.to_datetime(matched_df['EndDate'], errors='coerce')
    matched_df['Start Time'] = pd.to_datetime(matched_df['Start Time'], errors='coerce')
    matched_df['End Time'] = pd.to_datetime(matched_df['End Time'], errors='coerce')
    matched_df['Status'] = matched_df.apply(lambda row: 'Expired' if pd.notnull(row['End Time']) and row['End Time'] > row['EndDate'] else 'Valid', axis=1)
    return matched_df

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

# Automatically fetch the "User Name" file
user_name_file_path = "/path/to/USER NAME.xlsx"
user_name_df = pd.read_excel(user_name_file_path)

# Upload RMS and Current Alarms files
site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Alarms Excel", type=["xlsx"])

if site_access_file and rms_file and current_alarms_file:
    # Read files
    site_access_df = pd.read_excel(site_access_file)
    rms_df = pd.read_excel(rms_file, header=2)
    current_alarms_df = pd.read_excel(current_alarms_file, header=2)

    # Merge RMS and alarms data
    merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

    # Create left panel for filters and actions
    with st.sidebar:
        st.header("Filters and Actions")

        # Time filter
        selected_date = st.date_input("Select Date", value=datetime.now().date())
        selected_time = st.time_input("Select Time", value=datetime.now().time())
        filter_datetime = datetime.combine(selected_date, selected_time)

        # Dropdown for zone selection
        zones = site_access_df['Zone'].unique()
        selected_zone = st.selectbox("Select Zone to Modify Name", options=zones)
        if selected_zone:
            new_name = st.text_input(f"Enter new name for {selected_zone}")
            if st.button("Update Zone Name"):
                site_access_df.loc[site_access_df['Zone'] == selected_zone, 'Zone'] = new_name
                st.success(f"Zone name updated to {new_name}")

        # Button to send Telegram notifications
        if st.button("Send Telegram Notification"):
            mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
            zones = mismatches_df['Zone'].unique()
            bot_token = "7145427044:AAGb-CcT8zF_XYkutnqqCdNLqf6qw4KgqME"
            chat_id = "-4537588687"

            for zone in zones:
                zone_df = mismatches_df[mismatches_df['Zone'] == zone]
                message = f"*Door Open Notification*\n\n*{zone}*\n\n"
                site_aliases = zone_df['Site Alias'].unique()
                for site_alias in site_aliases:
                    site_df = zone_df[zone_df['Site Alias'] == site_alias]
                    message += f"#{site_alias}\n"
                    for _, row in site_df.iterrows():
                        end_time_display = row['End Time'] if row['End Time'] != 'Not Closed' else 'Not Closed'
                        message += f"Start Time: {row['Start Time']} End Time: {end_time_display}\n"
                    message += "\n"
                if send_telegram_notification(message, bot_token, chat_id):
                    st.success(f"Notification for zone '{zone}' sent successfully!")
                else:
                    st.error(f"Failed to send notification for zone '{zone}'.")

    # Process mismatches
    mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
    mismatches_df['Start Time'] = pd.to_datetime(mismatches_df['Start Time'], errors='coerce')
    filtered_mismatches_df = mismatches_df[mismatches_df['Start Time'] > filter_datetime]

    # Display mismatches
    st.write(f"Mismatched Sites (After {filter_datetime}):")
    st.dataframe(filtered_mismatches_df)

    # Process and display matches
    matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)
    st.write("Matched Sites:")
    st.dataframe(matched_df)

else:
    st.warning("Please upload all required files.")
