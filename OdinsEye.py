import streamlit as st
import pandas as pd
import re
from datetime import datetime

# Function to extract site name from 'SiteName' in Site Access Portal Log
def extract_site_name(site_name):
    """
    Extracts the site name from the Site Access Portal Log's SiteName column.
    It looks for patterns like 'NGBND85' or 'KHU0628' within the string.
    """
    # Regular expression to match patterns like NGBND85 or KHU0628
    match = re.search(r'[A-Z]{2,4}_?X?\d{4}', site_name)
    if match:
        return match.group()
    else:
        return None

# Function to extract site alias from 'Site Alias' in RMS Door Open Log
def extract_site_alias(site_alias):
    """
    Extracts the site alias from the RMS Door Open Log's Site Alias column.
    It assumes the site alias is before any space or parenthesis.
    """
    return site_alias.split(' ')[0].split('(')[0]

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
            rms_site_names = rms_df['Site Alias Clean'].unique()

            # Read Site Access Portal Log
            portal_df = pd.read_excel(portal_file)
            if 'SiteName' not in portal_df.columns:
                st.error("Site Access Portal Log must contain the 'SiteName' column.")
                return

            # Extract site names from Site Access Portal Log
            portal_df['Extracted SiteName'] = portal_df['SiteName'].apply(extract_site_name)
            portal_site_names = portal_df['Extracted SiteName'].dropna().unique()

            # Find unmatched site names
            unmatched_site_names = [site for site in rms_site_names if site not in portal_site_names]

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

            st.header("🔍 Unmatched Site Logs")
            st.write(f"Total Unmatched Sites: **{len(unmatched_site_names)}**")
            st.dataframe(display_df)

            # Filters
            st.sidebar.header("Filters")

            # Zone Filter
            zones = sorted(display_df['Zone'].unique())
            selected_zone = st.sidebar.selectbox("Select Zone", ["All"] + zones)

            # Cluster Filter
            clusters = sorted(display_df['Cluster'].unique())
            selected_cluster = st.sidebar.selectbox("Select Cluster", ["All"] + clusters)

            # Date Range Filter
            min_date = display_df['Start Time'].min().date()
            max_date = display_df['Start Time'].max().date()
            selected_date = st.sidebar.date_input("Select Date Range", [min_date, max_date])

            if len(selected_date) != 2:
                st.error("Please select a start and end date for the date range.")
                return

            start_datetime = datetime.combine(selected_date[0], datetime.min.time())
            end_datetime = datetime.combine(selected_date[1], datetime.max.time())

            # Apply Filters
            filtered_df = display_df.copy()

            if selected_zone != "All":
                filtered_df = filtered_df[filtered_df['Zone'] == selected_zone]

            if selected_cluster != "All":
                filtered_df = filtered_df[filtered_df['Cluster'] == selected_cluster]

            filtered_df = filtered_df[
                (filtered_df['Start Time'] >= start_datetime) &
                (filtered_df['Start Time'] <= end_datetime)
            ]

            st.header("📋 Filtered Unmatched Site Logs")
            st.write(f"Showing **{len(filtered_df)}** records after applying filters.")
            st.dataframe(filtered_df)

        except Exception as e:
            st.error(f"An error occurred while processing the files: {e}")

if __name__ == "__main__":
    main()
