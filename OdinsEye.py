import pandas as pd
import streamlit as st

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to find mismatches between SiteName from Site Access and Site from merged RMS/Alarms data
def find_mismatches(site_access_df, merged_df):
    # Extract the first part of SiteName for comparison
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)

    # Merge the data on SiteName_Extracted and Site to find mismatches
    merged_df = pd.merge(site_access_df, merged_df, left_on='SiteName_Extracted', right_on='Site', how='right', indicator=True)

    # Filter mismatched data (_merge column will have 'right_only' for missing entries in Site Access)
    mismatches_df = merged_df[merged_df['_merge'] == 'right_only']

    # Group by Cluster, Zone, Site Alias, Start Time, and End Time
    grouped_df = mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias', 'Start Time', 'End Time']).size().reset_index(name='Count')

    return grouped_df

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

    # Display the merged dataset
    st.write("Merged RMS and Current Alarms Dataset:")
    st.dataframe(merged_rms_alarms_df)  # Show the merged data in a table format

    # Find mismatches
    mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)

    if not mismatches_df.empty:
        st.write("Mismatched Sites grouped by Cluster and Zone:")
        display_grouped_data(mismatches_df, "Mismatched Sites")
    else:
        st.write("No mismatches found. All sites match between Site Access and RMS/Alarms.")
