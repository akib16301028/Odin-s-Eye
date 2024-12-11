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

# Function to display matched sites with status
def display_matched_sites(matched_df):
    color_map = {'Valid': 'background-color: lightgreen;', 'Expired': 'background-color: lightcoral;'}
    def highlight_status(status):
        return color_map.get(status, '')

    styled_df = matched_df[['RequestId', 'Site Alias', 'Start Time', 'End Time', 'EndDate', 'Status']].style.applymap(highlight_status, subset=['Status'])
    st.write("Matched Sites with Status:")
    st.dataframe(styled_df)

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

# Add a sidebar to display the USER NAME file
st.sidebar.title("File Information")

try:
    user_name_df = pd.read_csv("USER NAME.csv")  # Adjust path as needed
    st.sidebar.write("**USER NAME File Content:**")
    st.sidebar.table(user_name_df)
except FileNotFoundError:
    st.sidebar.error("USER NAME file not found. Ensure it exists in the repository.")

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

    # Move the "Send Telegram Notification" button to the top
    if st.button("Send Telegram Notification"):
        zones = filtered_mismatches_df['Zone'].unique()
        bot_token = "7145427044:AAGb-CcT8zF_XYkutnqqCdNLqf6qw4KgqME"
        chat_id = "-1001509039244"

        for zone in zones:
            zone_df = filtered_mismatches_df[filtered_mismatches_df['Zone'] == zone]
            message = f"*Door Open Alert*\n\n*{zone}*\n\n"  # Bold "Door Open Notification"
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
