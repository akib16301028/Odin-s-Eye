import pandas as pd
import streamlit as st

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to find mismatches between SiteName from Site Access and Site from RMS
def find_mismatches(site_access_df, rms_df):
    # Extract the first part of SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge the data on SiteName_Extracted and Site to find mismatches
    merged_df = pd.merge(site_access_df, rms_df, left_on='SiteName_Extracted', right_on='Site', how='right', indicator=True)

    # Filter mismatched data (_merge column will have 'right_only' for missing entries in Site Access)
    mismatches_df = merged_df[merged_df['_merge'] == 'right_only']

    # Group by Cluster, Zone, Site Alias
    grouped_df = mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias']).size().reset_index(name='Count')

    return grouped_df

# Function to find matched sites and their status
def find_matched_sites(site_access_df, rms_df):
    # Extract the first part of SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge the data on SiteName_Extracted and Site to find matches
    matched_df = pd.merge(site_access_df, rms_df, left_on='SiteName_Extracted', right_on='Site', how='inner')

    # Convert date columns to datetime
    matched_df['StartDate'] = pd.to_datetime(matched_df['StartDate'], errors='coerce')
    matched_df['EndDate'] = pd.to_datetime(matched_df['EndDate'], errors='coerce')
    matched_df['Start Time'] = pd.to_datetime(matched_df['Start Time'], errors='coerce')
    matched_df['End Time'] = pd.to_datetime(matched_df['End Time'], errors='coerce')

    # Identify valid and expired sites
    matched_df['Status'] = matched_df.apply(
        lambda row: 'Expired' if row['End Time'] > row['EndDate'] else 'Valid', axis=1
    )

    return matched_df

# Function to find mismatches in the current alarms
def find_alarm_mismatches(site_access_df, alarm_df):
    # Extract the first part of SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge the data on SiteName_Extracted and Site to find mismatches
    merged_df = pd.merge(site_access_df, alarm_df, left_on='SiteName_Extracted', right_on='Site', how='right', indicator=True)

    # Filter mismatched data (_merge column will have 'right_only' for missing entries in Site Access)
    alarm_mismatches_df = merged_df[merged_df['_merge'] == 'right_only']

    # Group by Cluster and Zone
    grouped_alarm_mismatches = alarm_mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias']).size().reset_index(name='Count')

    return grouped_alarm_mismatches

# Function to display grouped data by Cluster and Zone in a table with customized formatting
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
            display_df = zone_df[['Site Alias', 'Count']].copy()
            st.table(display_df)
        st.markdown("---")  # Separator between clusters

# Function to display matched sites with status
def display_matched_sites(matched_df):
    # Define colors based on status
    color_map = {'Valid': 'background-color: lightgreen;', 'Expired': 'background-color: lightcoral;'}
    
    def highlight_status(status):
        return color_map.get(status, '')

    # Apply the highlighting function
    styled_df = matched_df[['RequestId', 'Site Alias', 'Start Time', 'End Time', 'EndDate', 'Status']].style.applymap(highlight_status, subset=['Status'])

    st.write("Matched Sites:")
    st.dataframe(styled_df)

# Function to display grouped alarm mismatch data by Cluster and Zone
def display_grouped_alarm_mismatches(alarm_mismatches_df):
    st.write("Mismatched Alarms:")
    clusters = alarm_mismatches_df['Cluster'].unique()

    for cluster in clusters:
        st.markdown(f"**{cluster}**")  # Cluster in bold
        cluster_df = alarm_mismatches_df[alarm_mismatches_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"***<span style='font-size:14px;'>{zone}</span>***", unsafe_allow_html=True)  # Zone in italic bold with smaller font
            zone_df = cluster_df[cluster_df['Zone'] == zone]

            # Display only the relevant columns: Site Alias, Alarm Time (renamed to Start Time), Duration, and Duration Slot
            display_df = zone_df[['Site Alias', 'Alarm Time', 'Duration', 'Duration Slot (Hours)']].copy()
            display_df.rename(columns={'Alarm Time': 'Start Time'}, inplace=True)

            # Show the table
            st.table(display_df)
        st.markdown("---")  # Separator between clusters

# Streamlit app
st.title('Site Access and RMS Comparison Tool')

# File upload section
site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])
alarm_file = st.file_uploader("Upload the Current Alarms Excel", type=["xlsx"])

if site_access_file and rms_file:
    # Load the Site Access Excel
    site_access_df = pd.read_excel(site_access_file)

    # Load the RMS Excel, skipping the first two rows (so headers start from row 3)
    rms_df = pd.read_excel(rms_file, header=2)  # header=2 means row 3 is the header

    # Check if the necessary columns exist in both dataframes
    if 'SiteName' in site_access_df.columns and 'Site' in rms_df.columns:
        # Find mismatches
        mismatches_df = find_mismatches(site_access_df, rms_df)

        if not mismatches_df.empty:
            st.write("Mismatched Sites grouped by Cluster and Zone:")
            display_grouped_data(mismatches_df, "Mismatched Sites")
        else:
            st.write("No mismatches found. All sites match between Site Access and RMS.")

        # Find matched sites
        matched_df = find_matched_sites(site_access_df, rms_df)

        if not matched_df.empty:
            display_matched_sites(matched_df)
        else:
            st.write("No matched sites found.")

# After uploading the current alarms file
if alarm_file:
    # Load the Current Alarms Excel, skipping the first two rows (so headers start from row 3)
    alarm_df = pd.read_excel(alarm_file, header=2)  # header=2 means row 3 is the header

    # Check if the necessary columns exist in both dataframes
    if 'Site' in alarm_df.columns and 'SiteName' in site_access_df.columns:
        # Find mismatches in the current alarms
        alarm_mismatches_df = find_alarm_mismatches(site_access_df, alarm_df)

        if not alarm_mismatches_df.empty:
            display_grouped_alarm_mismatches(alarm_mismatches_df)
        else:
            st.write("No mismatches found in the Current Alarms file.")
    else:
        st.error("One or both files are missing the required columns (Site or SiteName).")
