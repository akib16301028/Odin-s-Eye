import pandas as pd
import streamlit as st
from datetime import datetime

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to standardize date and time formats for comparison
def standardize_datetime(date_series, time_series=None):
    if time_series is not None:
        # Combine date and time, if provided separately
        datetime_series = date_series + " " + time_series
    else:
        datetime_series = date_series
    
    # Convert to common datetime format
    return pd.to_datetime(datetime_series, errors='coerce')

# Function to find mismatches between SiteName from Site Access and Site from RMS
def find_mismatches(site_access_df, rms_df):
    # Extract the first part of SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge the data on SiteName_Extracted and Site to find mismatches
    merged_df = pd.merge(site_access_df, rms_df, left_on='SiteName_Extracted', right_on='Site', how='right', indicator=True)

    # Filter mismatched data (_merge column will have 'right_only' for missing entries in Site Access)
    mismatches_df = merged_df[merged_df['_merge'] == 'right_only']

    # Now check for entries that matched, but have an expired End Time
    matched_df = merged_df[merged_df['_merge'] == 'both']
    
    # Standardize date formats for comparison
    matched_df['SiteAccess_Start'] = standardize_datetime(matched_df['StartDate'])
    matched_df['SiteAccess_End'] = standardize_datetime(matched_df['EndDate'])
    matched_df['RMS_Start'] = standardize_datetime(matched_df['Start Time'])
    matched_df['RMS_End'] = standardize_datetime(matched_df['End Time'])
    
    # Check for expired entries where RMS End Time is greater than Site Access End Date
    expired_df = matched_df[matched_df['RMS_End'] > matched_df['SiteAccess_End']]
    expired_df['Expired'] = "Expired"

    # Combine mismatches and expired entries
    combined_df = pd.concat([mismatches_df, expired_df])
    
    # Group by Cluster, Zone, Site Alias, Start Time, and End Time, with 'Expired' column for expired entries
    grouped_df = combined_df.groupby(['Cluster', 'Zone', 'Site Alias', 'Start Time', 'End Time', 'Expired']).size().reset_index(name='Count')
    
    return grouped_df

# Function to display grouped data by Cluster and Zone in a table with customized formatting
def display_grouped_data(grouped_df):
    # First, group by Cluster, then by Zone
    clusters = grouped_df['Cluster'].unique()
    
    for cluster in clusters:
        # Display Cluster in bold
        st.markdown(f"**{cluster}**")  # Cluster in bold
        
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            # Display Zone in italic bold and slightly smaller
            st.markdown(f"***<span style='font-size:14px;'>{zone}</span>***", unsafe_allow_html=True)  # Zone in italic bold with smaller font
            
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            
            # Create a copy of the dataframe to handle hiding repeated Site Alias
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time', 'Expired']].copy()
            
            # Hide repeated Site Alias by replacing repeated values with empty strings
            display_df['Site Alias'] = display_df['Site Alias'].where(display_df['Site Alias'] != display_df['Site Alias'].shift())
            
            # Replace NaN values with empty strings to avoid <NA> display
            display_df = display_df.fillna('')
            
            # Highlight 'Expired' cells with a light pink background
            def highlight_expired(val):
                color = 'background-color: lightpink' if val == "Expired" else ''
                return [color] * len(val)

            # Display the table with styling
            st.dataframe(display_df.style.applymap(highlight_expired, subset=['Expired']))
        
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
        # Find mismatches and expired entries
        mismatches_df = find_mismatches(site_access_df, rms_df)

        if not mismatches_df.empty:
            st.write("Mismatched Sites and Expired entries grouped by Cluster and Zone:")
            display_grouped_data(mismatches_df)
        else:
            st.write("No mismatches found. All sites match between Site Access and RMS.")
    else:
        st.error("One or both files are missing the required columns (SiteName or Site).")
