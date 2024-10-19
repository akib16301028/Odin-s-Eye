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

    # Group by Cluster, Zone, Site Alias, Start Time, and End Time
    grouped_df = mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias', 'Start Time', 'End Time']).size().reset_index(name='Count')
    
    return grouped_df

# Function to display grouped data by Cluster and Zone in table format
def display_grouped_data(grouped_df):
    # First, group by Cluster, then by Zone
    clusters = grouped_df['Cluster'].unique()
    
    for cluster in clusters:
        st.subheader(f"{cluster}")
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"###{zone}")
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            
            # Create a new dataframe for display containing only the relevant columns
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']]
            
            # Display the table
            st.table(display_df)
        st.markdown("---")

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
            display_grouped_data(mismatches_df)
        else:
            st.write("No mismatches found. All sites match between Site Access and RMS.")
    else:
        st.error("One or both files are missing the required columns (SiteName or Site).")
