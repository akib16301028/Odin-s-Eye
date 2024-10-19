import pandas as pd
import streamlit as st

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to find mismatches between SiteName from Site Access and Site from RMS
def find_mismatches(site_access_df, rms_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_df = pd.merge(site_access_df, rms_df, left_on='SiteName_Extracted', right_on='Site', how='right', indicator=True)
    mismatches_df = merged_df[merged_df['_merge'] == 'right_only']
    grouped_df = mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias', 'Start Time', 'End Time']).size().reset_index(name='Count')
    return grouped_df

# Function to find mismatches with current alarms
def find_alarm_mismatches(site_access_df, alarms_df):
    # Extract the string before the underscore from SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    alarms_df['Site_Extracted'] = alarms_df['Site'].apply(extract_site)

    # Merge on the extracted columns to find mismatches
    merged_alarms = pd.merge(site_access_df, alarms_df, left_on='SiteName_Extracted', right_on='Site_Extracted', how='right', indicator=True)
    alarm_mismatches_df = merged_alarms[merged_alarms['_merge'] == 'right_only']

    # Return relevant columns, renaming Alarm Time to Start Time
    return alarm_mismatches_df[['RMS Station', 'Site', 'Site Alias', 'Zone', 'Cluster', 'Alarm Name', 'Tag', 'Tenant', 'Alarm Time', 'Duration', 'Duration Slot (Hours)']].rename(columns={'Alarm Time': 'Start Time'})

# Function to find matched sites and their status
def find_matched_sites(site_access_df, rms_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    matched_df = pd.merge(site_access_df, rms_df, left_on='SiteName_Extracted', right_on='Site', how='inner')
    matched_df['StartDate'] = pd.to_datetime(matched_df['StartDate'], errors='coerce')
    matched_df['EndDate'] = pd.to_datetime(matched_df['EndDate'], errors='coerce')
    matched_df['Start Time'] = pd.to_datetime(matched_df['Start Time'], errors='coerce')
    matched_df['End Time'] = pd.to_datetime(matched_df['End Time'], errors='coerce')
    matched_df['Status'] = matched_df.apply(lambda row: 'Expired' if row['End Time'] > row['EndDate'] else 'Valid', axis=1)
    return matched_df

# Function to display grouped data by Cluster and Zone
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
    st.write("Matched Sites:")
    st.dataframe(styled_df)

# Streamlit app
st.title('Site Access and RMS Comparison Tool')

# File upload section
site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])
alarms_file = st.file_uploader("Upload the Current Alarms Excel", type=["xlsx"])

if site_access_file and rms_file and alarms_file:
    # Load the Site Access Excel
    site_access_df = pd.read_excel(site_access_file)

    # Load the RMS Excel, skipping the first two rows (so headers start from row 3)
    rms_df = pd.read_excel(rms_file, header=2)

    # Load the Current Alarms Excel, skipping the first two rows (so headers start from row 3)
    alarms_df = pd.read_excel(alarms_file, header=2)

    # Check if the necessary columns exist in both dataframes
    if 'SiteName' in site_access_df.columns and 'Site' in rms_df.columns:
        # Find mismatches with RMS
        mismatches_df = find_mismatches(site_access_df, rms_df)

        if not mismatches_df.empty:
            st.write("Mismatched Sites grouped by Cluster and Zone:")
            display_grouped_data(mismatches_df, "Mismatched Sites")
        else:
            st.write("No mismatches found with RMS. All sites match between Site Access and RMS.")

        # Find matched sites
        matched_df = find_matched_sites(site_access_df, rms_df)

        if not matched_df.empty:
            display_matched_sites(matched_df)
        else:
            st.write("No matched sites found.")

        # Find mismatches with Current Alarms
        alarm_mismatches_df = find_alarm_mismatches(site_access_df, alarms_df)

        if not alarm_mismatches_df.empty:
            # Group by Cluster and Zone for display
            grouped_alarm_mismatches = alarm_mismatches_df.groupby(['Cluster', 'Zone']).size().reset_index(name='Count')
            st.write("Mismatched Alarms grouped by Cluster and Zone:")
            display_grouped_data(grouped_alarm_mismatches, "Mismatched Alarms")
        else:
            st.write("No mismatches found in the Current Alarms file.")
    else:
        st.error("One or both files are missing the required columns (SiteName or Site).")
