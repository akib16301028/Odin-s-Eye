import pandas as pd
import streamlit as st
from datetime import datetime
import requests

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
    matched_df['Status'] = matched_df.apply(
        lambda row: 'Expired' if pd.notnull(row['End Time']) and row['End Time'] > row['EndDate'] else 'Valid', axis=1)
    return matched_df

# Function to display grouped data by Cluster and Zone in a table
def display_grouped_data(grouped_df, title):
    st.write(title)
    clusters = grouped_df['Cluster'].unique()

    for cluster in clusters:
        st.markdown(f"**{cluster}**")
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"***<span style='font-size:14px;'>{zone}</span>***", unsafe_allow_html=True)
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']].copy()
            display_df['Site Alias'] = display_df['Site Alias'].where(display_df['Site Alias'] != display_df['Site Alias'].shift())
            display_df = display_df.fillna('')
            st.table(display_df)
        st.markdown("---")

# Function to send Telegram notification
def send_telegram_notification(message, bot_token, chat_id, parse_mode):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": parse_mode  # Change parse mode here (e.g., HTML, None)
    }
    response = requests.post(url, json=payload)
    return response.status_code == 200

# Streamlit app
st.title('Odin-s-Eye')

# Sidebar for upload and filter options
with st.sidebar:
    st.header("Upload Files")
    site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
    rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])
    current_alarms_file = st.file_uploader("Upload the Current Alarms Excel", type=["xlsx"])

    # Filters
    st.header("Filters")
    selected_date = st.date_input("Select Date", value=datetime.now().date())
    selected_time = st.time_input("Select Time", value=datetime.now().time())
    status_filter = st.selectbox("Filter by Status", options=["All", "Valid", "Expired"], index=0)

if site_access_file and rms_file and current_alarms_file:
    site_access_df = pd.read_excel(site_access_file)
    rms_df = pd.read_excel(rms_file, header=2)
    current_alarms_df = pd.read_excel(current_alarms_file, header=2)

    # Process data
    merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)
    mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
    matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)

    # Apply filters
    filter_datetime = datetime.combine(selected_date, selected_time)
    mismatches_df['Start Time'] = pd.to_datetime(mismatches_df['Start Time'], errors='coerce')
    filtered_mismatches_df = mismatches_df[mismatches_df['Start Time'] > filter_datetime]

    status_condition = (matched_df['Status'] == status_filter) if status_filter != "All" else True
    filtered_matched_df = matched_df[status_condition]

    # Main content area
    st.header("Results")
    st.subheader("Mismatched Sites")
    if not filtered_mismatches_df.empty:
        display_grouped_data(filtered_mismatches_df, "Filtered Mismatched Sites")
    else:
        st.write("No mismatches found for the selected filter criteria.")

    st.subheader("Matched Sites")
    st.dataframe(filtered_matched_df[['RequestId', 'Site Alias', 'Start Time', 'End Time', 'Status']])

    # Message generation and sending
    st.subheader("Message Notifications")
    if st.button("Generate and Send Messages"):
        bot_token = "7145427044:AAGb-CcT8zF_XYkutnqqCdNLqf6qw4KgqME"
        chat_id = "-4537588687"
        parse_mode = "HTML"  # Change parse mode as needed (e.g., Markdown, None)

        for zone, zone_df in filtered_mismatches_df.groupby('Zone'):
            message = f"<b>Door Open Notification</b>\n\n<b>{zone}</b>\n\n"
            for _, row in zone_df.iterrows():
                end_time = row['End Time'] if row['End Time'] != 'Not Closed' else "Not Closed"
                message += f"Start Time: {row['Start Time']}, End Time: {end_time}\n"
            message += "\nPlease address these alarms promptly."
            
            if send_telegram_notification(message, bot_token, chat_id, parse_mode):
                st.success(f"Message sent for Zone: {zone}")
            else:
                st.error(f"Failed to send message for Zone: {zone}")
else:
    st.write("Please upload all required files.")
