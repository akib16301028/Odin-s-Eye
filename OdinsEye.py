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

    # Group by Cluster, Zone, Site Alias, Alarm Time
    grouped_df = mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias', 'Alarm Time']).size().reset_index(name='Count')

    return grouped_df

# Function to find mismatches with Current Alarms
def find_current_alarm_mismatches(site_access_df, current_alarms_df):
    # Extract the first part of the Site from Current Alarms for comparison
    current_alarms_df['SiteName_Extracted'] = current_alarms_df['Site'].apply(extract_site)

    # Extract the SiteName from Site Access for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge to find mismatches
    merged_df = pd.merge(current_alarms_df, site_access_df, left_on='SiteName_Extracted', right_on='SiteName_Extracted', how='left', indicator=True)

    # Filter for mismatches
    mismatches_df = merged_df[merged_df['_merge'] == 'left_only']
    
    return mismatches_df

# Streamlit app
st.title('Site Access and RMS Comparison Tool')

# File upload section
site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Alarms Excel", type=["xlsx"])

if site_access_file and rms_file:
    # Load the Site Access Excel as-is
    site_access_df = pd.read_excel(site_access_file)

    # Load the RMS Excel, skipping the first two rows (so headers start from row 3)
    rms_df = pd.read_excel(rms_file, header=2)  # header=2 means row 3 is the header

    # Check if the necessary columns exist in both dataframes
    if 'SiteName' in site_access_df.columns and 'Site' in rms_df.columns:
        # Find mismatches
        mismatches_df = find_mismatches(site_access_df, rms_df)

        if not mismatches_df.empty:
            st.write("Mismatched Sites grouped by Cluster and Zone:")
            st.table(mismatches_df[['Cluster', 'Zone', 'Site Alias', 'Alarm Time']])  # Display required columns
        else:
            st.write("No mismatches found. All sites match between Site Access and RMS.")

        # After uploading Current Alarms file
        if current_alarms_file:
            # Load Current Alarms Excel, skipping the first two rows (so headers start from row 3)
            current_alarms_df = pd.read_excel(current_alarms_file, header=2)  # header=2 means row 3 is the header

            # Check if the necessary columns exist in current alarms dataframe
            if 'Site' in current_alarms_df.columns:
                # Find mismatches with current alarms
                alarm_mismatches = find_current_alarm_mismatches(site_access_df, current_alarms_df)

                if not alarm_mismatches.empty:
                    st.write("Mismatched Sites with Current Alarms:")
                    st.table(alarm_mismatches[['Site Alias', 'Zone', 'Cluster', 'Alarm Time']])  # Display required columns
                else:
                    st.write("No mismatches found in Current Alarms.")
            else:
                st.error("The Current Alarms file is missing the required column 'Site'.")
    else:
        st.error("One or both files are missing the required columns (SiteName or Site).")
