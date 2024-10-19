import pandas as pd
import streamlit as st
from datetime import datetime

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to merge RMS and Current Alarms data
def merge_rms_alarms(rms_df, alarms_df):
    # Use Alarm Time as Start Time for Current Alarms data
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT  # No End Time in Current Alarms, set to NaT

    # Keep the relevant columns from both RMS and Current Alarms
    rms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
    alarms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']

    # Concatenate RMS and Current Alarms data
    merged_df = pd.concat([rms_df[rms_columns], alarms_df[alarms_columns]], ignore_index=True)

    return merged_df

# Function to find mismatches between Site Access and merged RMS/Alarms dataset
def find_mismatches(site_access_df, merged_df):
    # Extract the first part of SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge the Site Access data with the merged RMS/Alarms dataset
    merged_comparison_df = pd.merge(merged_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)

    # Filter mismatched data (_merge column will have 'left_only' for missing entries in Site Access)
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']

    # Ensure that data without End Time is included
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('')

    # Group by Cluster, Zone, Site Alias, Start Time, and End Time
    grouped_df = mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias', 'Start Time', 'End Time'], dropna=False).size().reset_index(name='Count')

    return grouped_df

# Function to find matched sites and their status
def find_matched_sites(site_access_df, merged_df):
    # Extract the first part of SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge the data on SiteName_Extracted and Site to find matches
    matched_df = pd.merge(site_access_df, merged_df, left_on='SiteName_Extracted', right_on='Site', how='inner')

    # Convert date columns to datetime
    matched_df['StartDate'] = pd.to_datetime(matched_df['StartDate'], errors='coerce')
    matched_df['EndDate'] = pd.to_datetime(matched_df['EndDate'], errors='coerce')
    matched_df['Start Time'] = pd.to_datetime(matched_df['Start Time'], errors='coerce')
    matched_df['End Time'] = pd.to_datetime(matched_df['End Time'], errors='coerce')

    # Identify valid and expired sites
    matched_df['Status'] = matched_df.apply(
        lambda row: 'Expired' if pd.notnull(row['End Time']) and row['End Time'] > row['EndDate'] else 'Valid', axis=1
    )

    return matched_df

# Function to display grouped data by Cluster and Zone in a table
def display_grouped_data(grouped_df, title):
    st.write(title)
    clusters = grouped_df['Cluster'].unique()

    for cluster in clusters:
        st.markdown(f"**{cluster}**")  # Cluster in bold
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"***<span style='font-size:14px;'>{zone}</span>***", unsafe_allow_html=True)  # Zone in italic bold with smaller font
            
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']].copy()
            
            display_df['Site Alias'] = display_df['Site Alias'].where(display_df['Site Alias'] != display_df['Site Alias'].shift())
            display_df = display_df.fillna('')  # Replace NaN with empty strings

            st.table(display_df)
        st.markdown("---")  # Separator between clusters

# Function to display matched sites with status
def display_matched_sites(matched_df, start_time_filter=None):
    # Define colors based on status
    color_map = {'Valid': 'background-color: lightgreen;', 'Expired': 'background-color: lightcoral;'}

    def highlight_status(status):
        return color_map.get(status, '')

    # Filter matched_df based on selected start time filter
    if start_time_filter:
        matched_df = matched_df[matched_df['Start Time'] >= start_time_filter]

    # Apply the highlighting function
    styled_df = matched_df[['RequestId', 'Site Alias', 'Start Time', 'End Time', 'EndDate', 'Status']].style.applymap(highlight_status, subset=['Status'])

    st.write("Matched Sites with Status:")
    st.dataframe(styled_df)

# Function to display mismatched sites with status
def display_mismatched_sites(mismatched_df, start_time_filter=None):
    # Filter mismatched_df based on selected start time filter
    if start_time_filter:
        mismatched_df = mismatched_df[mismatched_df['Start Time'] >= start_time_filter]

    # Display the filtered mismatched DataFrame
    st.write("Mismatched Sites with Filter:")
    st.dataframe(mismatched_df)

# Streamlit app
st.title('Site Access and RMS/Alarms Comparison Tool')

# File upload section
site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Alarms Excel", type=["xlsx"])

if site_access_file and rms_file and current_alarms_file:
    # Load the Site Access Excel as-is
    site_access_df = pd.read_excel(site_access_file)

    # Load the RMS Excel, skipping the first two rows (so headers start from row 3)
    rms_df = pd.read_excel(rms_file, header=2)  # header=2 means row 3 is the header

    # Load the Current Alarms Excel, skipping the first two rows (so headers start from row 3)
    current_alarms_df = pd.read_excel(current_alarms_file, header=2)  # header=2 means row 3 is the header

    # Merge RMS and Current Alarms data
    merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

    # Compare Site Access with the merged RMS/Alarms dataset and find mismatches
    mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)

    # Add date and time filter for mismatched sites
    now = datetime.now()
    date_filter = st.date_input("Filter Mismatched Sites by Start Date:", value=now)
    time_filter = st.time_input("Filter Mismatched Sites by Start Time:", value=now.time())

    # Combine date and time filters into a single datetime object
    start_time_filter = pd.to_datetime(f"{date_filter} {time_filter}")

    if not mismatches_df.empty:
        display_mismatched_sites(mismatches_df, start_time_filter)
    else:
        st.write("No mismatches found. All sites match between Site Access and RMS/Alarms.")

    # Find matched sites and display with Valid/Expired status
    matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)

    if not matched_df.empty:
        # Display matched sites with status based on selected filter
        display_matched_sites(matched_df, start_time_filter)
    else:
        st.write("No matched sites found.")
