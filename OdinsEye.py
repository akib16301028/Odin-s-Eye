import pandas as pd
import streamlit as st

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to merge RMS and Current Alarms data
def merge_rms_alarms(rms_df, alarms_df):
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT  # No End Time in Current Alarms, set to NaT

    rms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
    alarms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']

    merged_df = pd.concat([rms_df[rms_columns], alarms_df[alarms_columns]], ignore_index=True)
    return merged_df

# Function to find mismatches between Site Access and merged RMS/Alarms dataset
def find_mismatches(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_comparison_df = pd.merge(merged_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('')

    grouped_df = mismatches_df.groupby(['Cluster', 'Zone', 'Site Alias', 'Start Time', 'End Time'], dropna=False).size().reset_index(name='Count')
    return grouped_df

# Function to find matched sites and their status
def find_matched_sites(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    matched_df = pd.merge(site_access_df, merged_df, left_on='SiteName_Extracted', right_on='Site', how='inner')

    matched_df['StartDate'] = pd.to_datetime(matched_df['StartDate'], errors='coerce')
    matched_df['EndDate'] = pd.to_datetime(matched_df['EndDate'], errors='coerce')
    matched_df['Start Time'] = pd.to_datetime(matched_df['Start Time'], errors='coerce')
    matched_df['End Time'] = pd.to_datetime(matched_df['End Time'], errors='coerce')

    matched_df['Status'] = matched_df.apply(
        lambda row: 'Expired' if pd.notnull(row['End Time']) and row['End Time'] > row['EndDate'] else 'Valid', axis=1
    )
    return matched_df

# Function to apply custom styling to tables
def style_table(df, status_column=None):
    def highlight_na(val):
        if pd.isna(val):
            return 'background-color: lightblue'
        return ''

    def highlight_default(val):
        return 'background-color: lightgray'

    if status_column:
        def highlight_status(val):
            if val == 'Valid':
                return 'background-color: lightgreen'
            elif val == 'Expired':
                return 'background-color: lightcoral'
            return ''
        styled_df = df.style.applymap(highlight_na).applymap(highlight_status, subset=[status_column]).applymap(highlight_default, subset=df.columns.difference([status_column]))
    else:
        styled_df = df.style.applymap(highlight_na).applymap(highlight_default)
    
    return styled_df

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

            st.table(style_table(display_df).to_html(), unsafe_allow_html=True)  # Style and display table
        st.markdown("---")  # Separator between clusters

# Function to display matched sites with status and filters
def display_matched_sites(matched_df):
    # Filter for Valid/Expired
    status_filter = st.selectbox("Filter by Status", options=['All', 'Valid', 'Expired'])
    if status_filter != 'All':
        matched_df = matched_df[matched_df['Status'] == status_filter]

    # DateTime filter for Start Time
    selected_datetime = st.date_input("Filter Start Time after", pd.to_datetime('now'))
    matched_df = matched_df[matched_df['Start Time'] >= pd.to_datetime(selected_datetime)]

    # Display the filtered and styled table
    st.write("Matched Sites with Status:")
    st.dataframe(style_table(matched_df, status_column='Status'))

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
    rms_df = pd.read_excel(rms_file, header=2)

    # Load the Current Alarms Excel, skipping the first two rows (so headers start from row 3)
    current_alarms_df = pd.read_excel(current_alarms_file, header=2)

    # Merge RMS and Current Alarms data
    merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

    # Compare Site Access with the merged RMS/Alarms dataset and find mismatches
    mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)

    # Filter option for mismatched sites by Start Time
    st.write("### Mismatched Sites")
    selected_mismatch_datetime = st.date_input("Filter Mismatched Sites Start Time after", pd.to_datetime('now'))
    filtered_mismatches_df = mismatches_df[mismatches_df['Start Time'] >= pd.to_datetime(selected_mismatch_datetime)]

    if not filtered_mismatches_df.empty:
        st.write("Mismatched Sites grouped by Cluster and Zone:")
        display_grouped_data(filtered_mismatches_df, "Mismatched Sites")
    else:
        st.write("No mismatches found. All sites match between Site Access and RMS/Alarms.")

    # Find matched sites and display with Valid/Expired status and time filter
    matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)

    if not matched_df.empty:
        display_matched_sites(matched_df)
    else:
        st.write("No matched sites found.")
