import pandas as pd
import streamlit as st
from fuzzywuzzy import process

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

# Function to search for Site Alias in Site Access data and display relevant info
def search_site_alias(site_access_df, search_term):
    # Extract the first part of the SiteName (as already done in the mismatches)
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    
    # Perform a case-insensitive search for matches using fuzzy matching
    choices = site_access_df['SiteName_Extracted'].unique()
    matches = process.extract(search_term, choices, limit=10, scorer=process.extractOne)
    
    # Filter the dataframe to include only rows with matching SiteName_Extracted
    matching_rows = site_access_df[site_access_df['SiteName_Extracted'].isin([match[0] for match in matches])]
    
    return matching_rows

# Function to display search results
def display_search_results(matching_rows):
    if matching_rows.empty:
        st.write("No matches found.")
    else:
        st.write(f"Found {len(matching_rows)} match(es):")
        # Display the relevant columns from Site Access for each match
        display_df = matching_rows[['RequestId', 'SiteName', 'SiteAccessType', 'StartDate', 'EndDate', 'InTime', 'OutTime', 'AccessPurpose', 'VendorName', 'POCName']]
        st.table(display_df)

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
        
        # Part 1: Mismatches Section (Existing functionality)
        mismatches_df = find_mismatches(site_access_df, rms_df)
        if not mismatches_df.empty:
            st.write("Mismatched Sites grouped by Cluster and Zone:")
            display_grouped_data(mismatches_df)
        else:
            st.write("No mismatches found. All sites match between Site Access and RMS.")
        
        # Part 2: Search Functionality
        st.write("### Search for Site Alias")
        search_term = st.text_input("Enter the Site Alias to search:")
        
        if search_term:
            # Perform the search
            search_results = search_site_alias(site_access_df, search_term)
            display_search_results(search_results)
        
    else:
        st.error("One or both files are missing the required columns (SiteName or Site).")
