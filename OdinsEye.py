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
    alarms_df['End Time'] = pd.NaT

    rms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
    alarms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']

    merged_df = pd.concat([rms_df[rms_columns], alarms_df[alarms_columns]], ignore_index=True)
    return merged_df

# Function to find mismatches
def find_mismatches(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_comparison_df = pd.merge(
        merged_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True
    )
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('Not Closed')
    return mismatches_df

# Function to display grouped data
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
def send_telegram_notification(message, bot_token, chat_id):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {"chat_id": chat_id, "text": message, "parse_mode": "Markdown"}
    response = requests.post(url, json=payload)
    return response.status_code == 200

# Load USER NAME table
try:
    user_name_df = pd.read_excel("USER NAME.xlsx")
except FileNotFoundError:
    user_name_df = pd.DataFrame(columns=['Zone', 'Concern Name'])

# Streamlit App
st.title("Odin-s-Eye")

# Sidebar Options
with st.sidebar:
    st.header("Options")
    show_user_table = st.checkbox("Show USER NAME Table")

    if show_user_table:
        st.write("USER NAME Table:")
        st.table(user_name_df)

    edit_user_table = st.checkbox("Edit USER NAME Table")
    if edit_user_table:
        edited_table = st.experimental_data_editor(user_name_df, num_rows="dynamic")
        if st.button("Save USER NAME Table"):
            edited_table.to_excel("USER NAME.xlsx", index=False)
            st.success("USER NAME Table updated!")

    send_zone_notification = st.checkbox("Send Zone-Wise Telegram Notification")
    if send_zone_notification:
        selected_zone = st.selectbox("Select Zone", user_name_df['Zone'].unique())
        concern_name = user_name_df[user_name_df['Zone'] == selected_zone]['Concern Name'].iloc[0]
        template_message = (
            f"@{concern_name}, no site access request found for following Door Open alarms. "
            "Please take care and share us update."
        )
        st.write("Message Template:")
        st.write(template_message)

        if st.button("Send Notification"):
            bot_token = "7145427044:AAGb-CcT8zF_XYkutnqqCdNLqf6qw4KgqME"
            chat_id = "-1001509039244"
            if send_telegram_notification(template_message, bot_token, chat_id):
                st.success("Notification sent successfully!")
            else:
                st.error("Failed to send notification.")

# File Upload
site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Alarms Excel", type=["xlsx"])

if site_access_file and rms_file and current_alarms_file:
    site_access_df = pd.read_excel(site_access_file)
    rms_df = pd.read_excel(rms_file, header=2)
    current_alarms_df = pd.read_excel(current_alarms_file, header=2)

    merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

    mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
    st.write("Mismatched Sites grouped by Cluster and Zone:")
    display_grouped_data(mismatches_df, "Mismatched Sites")
else:
    st.write("Please upload all required files.")
