import streamlit as st
import pandas as pd
from datetime import datetime

# Function to clean and extract site alias from RMS Door Open Log
def extract_site_alias(site_alias):
    """
    Extracts the site alias from the 'Site Alias' column.
    Assumes the format 'SITE123 (Description)' or similar.
    """
    if pd.isna(site_alias):
        return None
    return site_alias.split(' ')[0].split('(')[0].strip()

# Function to find matching site names in Site Access Portal Log
def find_matched_sites(portal_site_names, rms_site_names):
    """
    For each SiteName in the Site Access Portal Log, check if any RMS site name is a substring.
    Returns a set of matched RMS site names.
    """
    matched_sites = set()
    for portal_site in portal_site_names:
        if pd.isna(portal_site):
            continue
        for rms_site in rms_site_names:
            if rms_site in portal_site:
                matched_sites.add(rms_site)
    return matched_sites

# Streamlit App
def main():
    st.title("📊 Site Log Comparator")
    st.write("""
    This application compares site names from the **RMS Door Open Log** and the **Site Access Portal Log** to identify unmatched sites.
    """)

    st.sidebar.header("Upload Logs")

    # File uploaders
    rms_file = st.sidebar.file_uploader("Upload RMS Door Open Log (Excel)", type=["xlsx", "xls"])
    portal_file = st.sidebar.file_uploader("Upload Site Access Portal Log (Excel)", type=["xlsx", "xls"])

    if st.sidebar.button("Generate"):
        if not rms_file or not portal_file:
            st.error("Please upload both Excel files.")
            return

        try:
            # Read RMS Door Open Log
            rms_df = pd.read_excel(rms_file, skiprows=2)  # Skip first two rows
            required_rms_columns = ['Site Alias', 'Cluster', 'Zone', 'Start Time', 'End Time']
            if not all(col in rms_df.columns for col in required_rms_columns):
                st.error(f"RMS Door Open Log must contain the following columns: {required_rms_columns}")
                return

            # Clean and extract site aliases
            rms_df['Site Alias Clean'] = rms_df['Site Alias'].apply(extract_site_alias)
            rms_site_names = rms_df['Site Alias Clean'].dropna().unique()

            # Read Site Access Portal Log
            portal_df = pd.read_excel(portal_file)
            if 'SiteName' not in portal_df.columns:
                st.error("Site Access Portal Log must contain the 'SiteName' column.")
                return

            portal_site_names = portal_df['SiteName'].dropna().unique()

            # Find matched site names
            matched_site_names = find_matched_sites(portal_site_names, rms_site_names)

            # Find unmatched site names
            unmatched_site_names = [site for site in rms_site_names if site not in matched_site_names]

            if not unmatched_site_names:
                st.success("All site names from RMS Door Open Log are matched in Site Access Portal Log.")
                return

            # Filter RMS data for unmatched sites
            unmatched_df = rms_df[rms_df['Site Alias Clean'].isin(unmatched_site_names)].copy()

            # Convert 'Start Time' and 'End Time' to datetime
            unmatched_df['Start Time'] = pd.to_datetime(unmatched_df['Start Time'], errors='coerce')
            unmatched_df['End Time'] = pd.to_datetime(unmatched_df['End Time'], errors='coerce')

            # Drop rows with invalid dates
            unmatched_df = unmatched_df.dropna(subset=['Start Time', 'End Time'])

            # Select relevant columns
            display_df = unmatched_df[['Site Alias', 'Cluster', 'Zone', 'Start Time', 'End Time']].copy()

            # Rename columns for better readability
            display_df.rename(columns={
                'Site Alias': 'Site Alias',
                'Cluster': 'Cluster',
                'Zone': 'Zone',
                'Start Time': 'Start Time',
                'End Time': 'End Time'
            }, inplace=True)

            # Convert datetime to string for display
            display_df['Start Time'] = display_df['Start Time'].dt.strftime('%m/%d/%Y %I:%M:%S %p')
            display_df['End Time'] = display_df['End Time'].dt.strftime('%m/%d/%Y %I:%M:%S %p')

            st.header("🔍 Unmatched Site Logs")
            st.write(f"**Total Unmatched Sites: {len(unmatched_site_names)}**")
            st.dataframe(display_df)

            # Filters
            st.sidebar.header("Filters")

            # Zone Filter
            zones = sorted(display_df['Zone'].dropna().unique())
            selected_zone = st.sidebar.selectbox("Select Zone", ["All"] + list(zones))

            # Cluster Filter
            clusters = sorted(display_df['Cluster'].dropna().unique())
            selected_cluster = st.sidebar.selectbox("Select Cluster", ["All"] + list(clusters))

            # Date Range Filter
            # Convert 'Start Time' back to datetime for filtering
            unmatched_df['Start Time'] = pd.to_datetime(unmatched_df['Start Time'], errors='coerce')
            min_date = unmatched_df['Start Time'].min().date()
            max_date = unmatched_df['Start Time'].max().date()
            selected_date = st.sidebar.date_input("Select Date Range", [min_date, max_date])

            if len(selected_date) != 2:
                st.error("Please select a start and end date for the date range.")
                return

            start_datetime = datetime.combine(selected_date[0], datetime.min.time())
            end_datetime = datetime.combine(selected_date[1], datetime.max.time())

            # Apply Filters
            filtered_df = unmatched_df.copy()

            if selected_zone != "All":
                filtered_df = filtered_df[filtered_df['Zone'] == selected_zone]

            if selected_cluster != "All":
                filtered_df = filtered_df[filtered_df['Cluster'] == selected_cluster]

            filtered_df = filtered_df[
                (filtered_df['Start Time'] >= start_datetime) &
                (filtered_df['Start Time'] <= end_datetime)
            ]

            # Prepare data for display
            if not filtered_df.empty:
                filtered_display_df = filtered_df[['Site Alias', 'Cluster', 'Zone', 'Start Time', 'End Time']].copy()
                filtered_display_df['Start Time'] = filtered_display_df['Start Time'].dt.strftime('%m/%d/%Y %I:%M:%S %p')
                filtered_display_df['End Time'] = filtered_display_df['End Time'].dt.strftime('%m/%d/%Y %I:%M:%S %p')
            else:
                filtered_display_df = pd.DataFrame(columns=['Site Alias', 'Cluster', 'Zone', 'Start Time', 'End Time'])

            st.header("📋 Filtered Unmatched Site Logs")
            st.write(f"**Showing {len(filtered_display_df)} record(s) after applying filters.**")
            st.dataframe(filtered_display_df)

        except Exception as e:
            st.error(f"An error occurred while processing the files: {e}")

if __name__ == "__main__":
    main()
