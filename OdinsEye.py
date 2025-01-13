import pandas as pd
import streamlit as st
from datetime import datetime
import requests  # For Telegram notifications

# Function to extract site name (before first underscore)
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to load zone-to-name mapping
def load_zone_name_mapping(file_path="USER NAME.xlsx"):
    try:
        user_name_df = pd.read_excel(file_path)
        return dict(zip(user_name_df['Zone'], user_name_df['Name']))
    except Exception as e:
        st.error(f"Error loading USER NAME.xlsx: {e}")
        return {}

# Function to merge RMS and current alarms datasets
def merge_rms_alarms(rms_df, alarms_df):
    try:
        alarms_df['Start Time'] = alarms_df['Alarm Time']
        alarms_df['End Time'] = pd.NaT  # Current alarms have no End Time
        rms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
        alarms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
        
        if not set(rms_columns).issubset(rms_df.columns) or not set(alarms_columns).issubset(alarms_df.columns):
            raise ValueError("RMS or Alarms DataFrame is missing required columns.")
        
        merged_df = pd.concat([rms_df[rms_columns], alarms_df[alarms_columns]], ignore_index=True)
        return merged_df
    except Exception as e:
        st.error(f"Error merging datasets: {e}")
        return pd.DataFrame()

# Function to identify mismatches
def find_mismatches(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_comparison = pd.merge(merged_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    mismatches_df = merged_comparison[merged_comparison['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('Not Closed')
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

# Function to display grouped data
def display_grouped_data(grouped_df, title):
    st.write(title)
    for cluster in grouped_df['Cluster'].unique():
        st.markdown(f"**{cluster}**")
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        for zone in cluster_df['Zone'].unique():
            st.markdown(f"***{zone}***")
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']].fillna('')
            st.table(display_df)

# Function to send Telegram notifications
def send_telegram_notification(message, bot_token, chat_id):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {"chat_id": chat_id, "text": message, "parse_mode": "Markdown"}
    response = requests.post(url, json=payload)
    return response.status_code == 200

# Streamlit app starts here
st.title("Odin's Eye")

# File uploads
site_access_file = st.file_uploader("Upload Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload RMS Excel", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload Current Alarms Excel", type=["xlsx"])

if site_access_file and rms_file and current_alarms_file:
    try:
        site_access_df = pd.read_excel(site_access_file)
        rms_df = pd.read_excel(rms_file, header=2)
        current_alarms_df = pd.read_excel(current_alarms_file, header=2)

        # Parse date/time columns
        for df in [rms_df, current_alarms_df]:
            df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
            df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce')

        merged_df = merge_rms_alarms(rms_df, current_alarms_df)
        if not merged_df.empty:
            mismatches_df = find_mismatches(site_access_df, merged_df)
            matched_df = find_matched_sites(site_access_df, merged_df)

            # Display mismatches and matched sites
            st.write("Mismatched Sites:")
            st.dataframe(mismatches_df)

            st.write("Matched Sites:")
            st.dataframe(matched_df)

            # Telegram notification functionality
            if st.button("Send Telegram Notifications"):
                bot_token = "YOUR_BOT_TOKEN"
                chat_id = "YOUR_CHAT_ID"
                zones = mismatches_df['Zone'].unique()

                for zone in zones:
                    zone_df = mismatches_df[mismatches_df['Zone'] == zone]
                    message = f"*Zone Alert: {zone}*\n"
                    for _, row in zone_df.iterrows():
                        message += f"- Site: {row['Site Alias']}\n  Start Time: {row['Start Time']}\n  End Time: {row['End Time']}\n"
                    if send_telegram_notification(message, bot_token, chat_id):
                        st.success(f"Notification sent for Zone: {zone}")
                    else:
                        st.error(f"Failed to send notification for Zone: {zone}")
    except Exception as e:
        st.error(f"Error processing files: {e}")
