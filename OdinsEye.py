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
    
    return mismatches_df

# Function to identify expired sites based on date ranges
def find_expired_sites(site_access_df, rms_df):
    # Extract the first part of SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge the data on SiteName_Extracted and Site to find matches
    merged_df = pd.merge(site_access_df, rms_df, left_on='SiteName_Extracted', right_on='Site', how='inner')

    # Convert date columns to datetime
    merged_df['StartDate'] = pd.to_datetime(merged_df['StartDate'], errors='coerce')
    merged_df['EndDate'] = pd.to_datetime(merged_df['EndDate'], errors='coerce')
    merged_df['Start Time'] = pd.to_datetime(merged_df['Start Time'], errors='coerce')
    merged_df['End Time'] = pd.to_datetime(merged_df['End Time'], errors='coerce')

    # Identify expired sites where RMS End Time is greater than Site Access EndDate
    expired_df = merged_df[merged_df['End Time'] > merged_df['EndDate']]

    return expired_df

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
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']].copy()
            
            display_df['Site Alias'] = display_df['Site Alias'].where(display_df['Site Alias'] != display_df['Site Alias'].shift())
            display_df = display_df.fillna('')  # Replace NaN with empty strings

            st.table(display_df)
        st.markdown("---")  # Separator between clusters

# Streamlit app
st.title('Site Access and RMS Comparison Tool')

# File upload section
site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])

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
            display_grouped_data(mismatches_df, "Mismatched Sites")

        else:
            st.write("No mismatches found. All sites match between Site Access and RMS.")

        # Find expired sites
        expired_df = find_expired_sites(site_access_df, rms_df)

        if not expired_df.empty:
            display_grouped_data(expired_df, "Expired Sites")
        else:
            st.write("No expired sites found.")
    else:
        st.error("One or both files are missing the required columns (SiteName or Site).")
