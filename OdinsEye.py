import streamlit as st
import pandas as pd
from datetime import datetime

# Function to clean and extract site alias from RMS Door Open Log
def extract_site_alias(site_alias):
    """
    Extracts the site alias from the 'Site Alias' column.
    Assumes the format 'SITE123 (Description)' or similar.
    Example:
        'DHK3170 (BL)' -> 'DHK3170'
        'DHKU23 (ROBI)' -> 'DHKU23'
    """
    if pd.isna(site_alias):
        return None
    return site_alias.split(' ')[0].split('(')[0].strip().upper()

# Function to extract site name from Site Access Portal Log's SiteName column
def extract_site_name(portal_site_name):
    """
    Extracts the site name from the Site Access Portal Log's SiteName column.
    Rules:
        - If SiteName is 'ABCDEF_anystring', extract 'anystring'
        - If SiteName is 'ABCDEF_anystring_X1234' or 'ABCDEF_anystring_A1234', extract 'anystring_X1234' or 'anystring_A1234'
          (i.e., preserve the letter after the second underscore)
    Examples:
        'SBCMNGK001_DHK_X3170' -> 'DHK_X3170'
        'SBCMNGK001_DHKU23' -> 'DHKU23'
        'ABCDEF_anystring_A1234' -> 'ANYSTRING_A1234'
    """
    if pd.isna(portal_site_name):
        return None
    parts = portal_site_name.split('_')
    if len(parts) == 2:
        # Format: 'ABCDEF_anystring' -> 'anystring'
        return parts[1].strip().upper()
    elif len(parts) >= 3:
        # Format: 'ABCDEF_anystring_X1234' or 'ABCDEF_anystring_A1234' -> 'anystring_X1234' or 'anystring_A1234'
        site_part = parts[1].strip().upper()
        suffix_part = parts[2].strip()
        if len(suffix_part) >= 1 and suffix_part[0].isalpha():
            # Preserve the leading letter and concatenate
            return f"{site_part}_{suffix_part}".upper()
        else:
            # Concatenate without modification
            return f"{site_part}_{suffix_part}".upper()
    else:
        # Unexpected format, return the last part
        return parts[-1].strip().upper()

# Function to find matched site names
def find_matched_sites(portal_site_names, rms_site_names):
    """
    Compares site names from RMS and Site Access Portal logs.
    Returns a set of matched RMS site names.
    """
    matched_sites = set()
    portal_site_set = set(portal_site_names)
    for rms_site in rms_site_names:
        if rms_site in portal_site_set:
            matched_sites.add(rms_site)
    return matched_sites

# Streamlit App
def main():
    st.set_page_config(page_title="📊 Site Log Comparator", layout="wide")
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
            with st.spinner('Processing files...'):
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

                # Extract site names from Site Access Portal Log
                portal_df['Extracted SiteName'] = portal_df['SiteName'].apply(extract_site_name)
                portal_site_names = portal_df['Extracted SiteName'].dropna().unique()

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

                # Convert datetime to string for display
                display_df['Start Time'] = display_df['Start Time'].dt.strftime('%m/%d/%Y %I:%M:%S %p')
                display_df['End Time'] = display_df['End Time'].dt.strftime('%m/%d/%Y %I:%M:%S %p')

            # Display Unmatched Site Logs
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

            # Display Filtered Unmatched Site Logs
            st.header("📋 Filtered Unmatched Site Logs")
            st.write(f"**Showing {len(filtered_display_df)} record(s) after applying filters.**")
            st.dataframe(filtered_display_df)

        except Exception as e:
            st.error(f"An error occurred while processing the files: {e}")

if __name__ == "__main__":
    main()
