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
    grouped_df = mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias']).size().reset_index(name='Count')
    return grouped_df

# Function to find matched sites and their status
def find_matched_sites(site_access_df, rms_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    matched_df = pd.merge(site_access_df, rms_df, left_on='SiteName_Extracted', right_on='Site', how='inner')
    matched_df['Status'] = matched_df.apply(lambda row: 'Expired' if row['End Time'] > row['EndDate'] else 'Valid', axis=1)
    return matched_df

# Function to find mismatches in the current alarms
def find_alarm_mismatches(site_access_df, alarm_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_df = pd.merge(site_access_df, alarm_df, left_on='SiteName_Extracted', right_on='Site', how='right', indicator=True)
    alarm_mismatches_df = merged_df[merged_df['_merge'] == 'right_only']
    grouped_alarm_mismatches = alarm_mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias']).size().reset_index(name='Count')
    return grouped_alarm_mismatches

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

            # Check for the existence of relevant columns
            relevant_columns = ['Site Alias', 'Alarm Time', 'Duration', 'Duration Slot (Hours)']
            missing_columns = [col for col in relevant_columns if col not in zone_df.columns]

            if missing_columns:
                st.error(f"Missing columns in the data for {zone}: {', '.join(missing_columns)}")
                continue

            # Display the relevant columns
            display_df = zone_df[relevant_columns].copy()
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
    site_access_df = pd.read_excel(site_access_file)
    rms_df = pd.read_excel(rms_file, header=2)  # header=2 means row 3 is the header

    if 'SiteName' in site_access_df.columns and 'Site' in rms_df.columns:
        mismatches_df = find_mismatches(site_access_df, rms_df)

        if not mismatches_df.empty:
            st.write("Mismatched Sites grouped by Cluster and Zone:")
            display_grouped_data(mismatches_df, "Mismatched Sites")
        else:
            st.write("No mismatches found. All sites match between Site Access and RMS.")

        matched_df = find_matched_sites(site_access_df, rms_df)

        if not matched_df.empty:
            display_matched_sites(matched_df)
        else:
            st.write("No matched sites found.")

if alarm_file:
    alarm_df = pd.read_excel(alarm_file, header=2)  # header=2 means row 3 is the header

    if 'Site' in alarm_df.columns and 'SiteName' in site_access_df.columns:
        alarm_mismatches_df = find_alarm_mismatches(site_access_df, alarm_df)

        if not alarm_mismatches_df.empty:
            display_grouped_alarm_mismatches(alarm_mismatches_df)
        else:
            st.write("No mismatches found in the Current Alarms file.")
    else:
        st.error("One or both files are missing the required columns (Site or SiteName).")
