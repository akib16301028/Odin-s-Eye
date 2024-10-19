import pandas as pd
import streamlit as st

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
    merged_df = pd.concat([rms_df[rms_columns], alarms_df[alarms_columns]])

    return merged_df

# Function to find mismatches between Site Access and the merged RMS/Alarms data
def find_mismatches(site_access_df, merged_df):
    # Extract the first part of SiteName for comparison in Site Access
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Extract the first part of the Site for comparison in merged data
    merged_df['Site_Extracted'] = merged_df['Site'].apply(extract_site)

    # Merge the data on extracted site names
    merged_compare_df = pd.merge(site_access_df, merged_df, left_on='SiteName_Extracted', right_on='Site_Extracted', how='right', indicator=True)

    # Filter mismatches (entries that are only in the merged RMS/Alarms dataset)
    mismatches_df = merged_compare_df[merged_compare_df['_merge'] == 'right_only']

    # Group by Cluster, Zone, Site Alias, Start Time, and End Time
    grouped_df = mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias', 'Start Time', 'End Time']).size().reset_index(name='Count')

    return grouped_df

# Function to find matched sites and their status
def find_matched_sites(site_access_df, merged_df):
    # Extract the first part of SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge the data on SiteName_Extracted and Site_Extracted to find matches
    matched_df = pd.merge(site_access_df, merged_df, left_on='SiteName_Extracted', right_on='Site_Extracted', how='inner')

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

    # Check if the necessary columns exist in all dataframes
    if 'SiteName' in site_access_df.columns and 'Site' in merged_rms_alarms_df.columns:
        # Find mismatches
        mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)

        if not mismatches_df.empty:
            st.write("Mismatched Sites grouped by Cluster and Zone:")
            st.table(mismatches_df[['Cluster', 'Zone', 'Site Alias', 'Start Time', 'End Time']])  # Display required columns
        else:
            st.write("No mismatches found. All sites match between Site Access and RMS/Alarms.")

        # Find matched sites
        matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)

        if not matched_df.empty:
            st.write("Matched Sites with Status:")
            st.table(matched_df[['RequestId', 'Site Alias', 'Start Time', 'End Time', 'EndDate', 'Status']])  # Display matched sites with status
        else:
            st.write("No matched sites found.")
    else:
        st.error("One or more files are missing the required columns.")
