import pandas as pd
import streamlit as st
from datetime import datetime, time
import requests

# Function to preprocess current alarm files to match expected column names
def preprocess_current_alarm(df, alarm_type):
    df = df.rename(columns={"Alarm Time": "Start Time"})
    df["End Time"] = None  # Set End Time as None for current alarm entries
    df["Type"] = alarm_type  # Specify type as either 'Motion' or 'Vibration'
    return df

# Function to merge motion and vibration data from reports and current alarms
def merge_motion_vibration(report_motion_df, current_motion_df, report_vibration_df, current_vibration_df):
    report_motion_df['Type'] = 'Motion'
    report_vibration_df['Type'] = 'Vibration'
    
    for df in [report_motion_df, report_vibration_df]:
        df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
        df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce')

    current_motion_df = preprocess_current_alarm(current_motion_df, 'Motion')
    current_vibration_df = preprocess_current_alarm(current_vibration_df, 'Vibration')

    merged_df = pd.concat([report_motion_df, current_motion_df, report_vibration_df, current_vibration_df], ignore_index=True)
    return merged_df

# Function to count occurrences of Motion and Vibration events per Site Alias and Zone
def count_entries_by_zone(merged_df, start_time_filter=None):
    if start_time_filter:
        merged_df = merged_df[merged_df['Start Time'] > start_time_filter]

    motion_count = merged_df[merged_df['Type'] == 'Motion'].groupby(['Zone', 'Site Alias']).size().reset_index(name='Motion Count')
    vibration_count = merged_df[merged_df['Type'] == 'Vibration'].groupby(['Zone', 'Site Alias']).size().reset_index(name='Vibration Count')
    
    final_df = pd.merge(motion_count, vibration_count, on=['Zone', 'Site Alias'], how='outer').fillna(0)
    final_df['Motion Count'] = final_df['Motion Count'].astype(int)
    final_df['Vibration Count'] = final_df['Vibration Count'].astype(int)
    
    return final_df

# Function to display detailed entries for a specific site alias
def display_detailed_entries(merged_df, site_alias):
    filtered = merged_df[merged_df['Site Alias'] == site_alias][['Site Alias', 'Start Time', 'End Time', 'Type']]
    if not filtered.empty:
        st.sidebar.write(f"Detailed entries for Site Alias: {site_alias}")
        st.sidebar.table(filtered)
    else:
        st.sidebar.write("No data for this site.")

# Function to send Telegram notification
def send_telegram_notification(zone, zone_df, total_motion, total_vibration):
    chat_id = "-4537588687"
    bot_token = "7145427044:AAGb-CcT8zF_XYkutnqqCdNLqf6qw4KgqME"
    
    # Construct message
    message = f"**Zone: {zone}**\n\n"
    message += f"Total Motion Alarm count: {total_motion}\nTotal Vibration Alarm count: {total_vibration}\n\n"
    for _, row in zone_df.iterrows():
        message += f"**{row['Site Alias']}** : Motion Count: {row['Motion Count']}, Vibration Count: {row['Vibration Count']}\n"
    
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    data = {"chat_id": chat_id, "text": message, "parse_mode": "Markdown"}
    
    response = requests.post(url, data=data)
    if response.status_code == 200:
        st.success(f"Notification sent for Zone: {zone}")
    else:
        st.error("Failed to send notification.")

# Streamlit app
st.title('Odin-s-Eye - Motion & Vibration Alarm Monitoring')

# File upload section
report_motion_file = st.file_uploader("Upload the Motion Report Data", type=["xlsx"])
current_motion_file = st.file_uploader("Upload the Motion Current Alarms Data", type=["xlsx"])
report_vibration_file = st.file_uploader("Upload the Vibration Report Data", type=["xlsx"])
current_vibration_file = st.file_uploader("Upload the Vibration Current Alarms Data", type=["xlsx"])

if report_motion_file and current_motion_file and report_vibration_file and current_vibration_file:
    report_motion_df = pd.read_excel(report_motion_file, header=2)
    current_motion_df = pd.read_excel(current_motion_file, header=2)
    report_vibration_df = pd.read_excel(report_vibration_file, header=2)
    current_vibration_df = pd.read_excel(current_vibration_file, header=2)

    merged_df = merge_motion_vibration(report_motion_df, current_motion_df, report_vibration_df, current_vibration_df)

    # Sidebar options for download and notifications
    with st.sidebar:
        # Download report button
        csv_data = merged_df.to_csv(index=False).encode('utf-8')
        st.download_button(label="Download Report as CSV", data=csv_data, file_name="alarm_summary.csv", mime="text/csv")

        # Date and time filter
        selected_date = st.date_input("Select Start Date", value=datetime.now().date())
        selected_time = st.time_input("Select Start Time", value=time(0, 0))
        start_time_filter = datetime.combine(selected_date, selected_time)

        # Telegram notification button
        if st.button("Telegram Notification", help="Send alarm summary to Telegram"):
            summary_df = count_entries_by_zone(merged_df, start_time_filter)
            zones = summary_df['Zone'].unique()
            for zone in zones:
                zone_df = summary_df[summary_df['Zone'] == zone]
                total_motion = zone_df['Motion Count'].sum()
                total_vibration = zone_df['Vibration Count'].sum()
                send_telegram_notification(zone, zone_df, total_motion, total_vibration)

    summary_df = count_entries_by_zone(merged_df, start_time_filter)

    zones = summary_df['Zone'].unique()
    for zone in zones:
        st.write(f"### Zone: {zone}")
        zone_df = summary_df[summary_df['Zone'] == zone]

        # Calculate total counts for the zone
        total_motion = zone_df['Motion Count'].sum()
        total_vibration = zone_df['Vibration Count'].sum()

        # Display total counts
        st.write(f"Total Motion Alarm count: {total_motion}")
        st.write(f"Total Vibration Alarm count: {total_vibration}")

        # Display the detailed table
        st.table(zone_df[['Zone', 'Site Alias', 'Motion Count', 'Vibration Count']])

    site_search = st.sidebar.text_input("Search for a specific site alias")
    if site_search:
        display_detailed_entries(merged_df, site_search)
else:
    st.write("Please upload all required files.")
