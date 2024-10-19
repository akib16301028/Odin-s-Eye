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

# Function to display matched sites with status
def display_matched_sites(matched_df):
    color_map = {'Valid': 'background-color: lightgreen;', 'Expired': 'background-color: lightcoral;'}
    
    def highlight_status(status):
        return color_map.get(status, '')

    styled_df = matched_df[['RequestId', 'Site Alias', 'Start Time', 'End Time', 'EndDate', 'Status']].style.applymap(highlight_status, subset=['Status'])
    st.write("Matched Sites:")
    st.dataframe(styled_df)

# Function to find mismatches with Current Alarms
def find_mismatches_with_alarms(mismatches_df, alarms_df):
    # Merge mismatches with alarms based on Site Alias
    merged_alarms = pd.merge(mismatches_df, alarms_df, left_on='Site Alias', right_on='Site', how='left')
    return merged_alarms[['Cluster', 'Zone', 'Site Alias', 'Alarm Time', 'Start Time', 'End Time']]

# Streamlit app
st.title('Site Access and RMS Comparison Tool')

# File upload section
site_access_file = st.file_uploader("Upload the Site Access Excel", type=["xlsx"])
rms_file = st.file_uploader("Upload the RMS Excel", type=["xlsx"])
alarms_file = st.file_uploader("Upload the Current Alarms Excel", type=["xlsx"])

# Generate button
if st.button('Generate'):
    if site_access_file and rms_file and alarms_file:
        # Load the files
        site_access_df = pd.read_excel(site_access_file)
        rms_df = pd.read_excel(rms_file, header=2)  # header=2 means row 3 is the header
        alarms_df = pd.read_excel(alarms_file)

        # Check for required columns in both dataframes
        if 'SiteName' in site_access_df.columns and 'Site' in rms_df.columns and 'Site' in alarms_df.columns:
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

            # Find mismatches with Current Alarms
            merged_with_alarms = find_mismatches_with_alarms(mismatches_df, alarms_df)
            if not merged_with_alarms.empty:
                st.write("Mismatched Sites with Current Alarms:")
                st.table(merged_with_alarms)
            else:
                st.write("No mismatched sites found with Current Alarms.")
        else:
            st.error("One or more files are missing the required columns.")
    else:
        st.warning("Please upload all required files.")
