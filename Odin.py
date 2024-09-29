import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# ============================
# Alarm and Offline Monitoring
# ============================

# Function to extract client name from Site Alias
def extract_client(site_alias):
    match = re.search(r'\((.*?)\)', site_alias)
    return match.group(1) if match else None

# Function to create pivot table for a specific alarm
def create_pivot_table(df, alarm_name):
    alarm_df = df[df['Alarm Name'] == alarm_name].copy()

    if alarm_name == 'DCDB-01 Primary Disconnect':
        alarm_df = alarm_df[~alarm_df['RMS Station'].str.startswith('L')]

    pivot = pd.pivot_table(
        alarm_df,
        index=['Cluster', 'Zone'],
        columns='Client',
        values='Site Alias',
        aggfunc='count',
        fill_value=0
    )

    pivot = pivot.reset_index()
    client_columns = [col for col in pivot.columns if col not in ['Cluster', 'Zone']]
    pivot['Total'] = pivot[client_columns].sum(axis=1)

    total_row = pivot[client_columns + ['Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']

    pivot = pd.concat([pivot, total_row], ignore_index=True)

    total_alarm_count = pivot['Total'].iloc[-1]

    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']

    return pivot, total_alarm_count

# Function to create pivot table for offline report
def create_offline_pivot(df):
    df = df.drop_duplicates()

    df['1 or Less than 1 day'] = df['Duration'].apply(
        lambda x: 1 if x in ['-1', '0', '1', 'Less than 24 hours'] else 0
    )
    df['More than 24 hours'] = df['Duration'].apply(
        lambda x: 1 if 'More than 24 hours' in x and '72' not in x else 0
    )
    df['More than 72 hours'] = df['Duration'].apply(
        lambda x: 1 if 'More than 72 hours' in x else 0
    )

    pivot = df.groupby(['Cluster', 'Zone']).agg({
        '1 or Less than 1 day': 'sum',
        'More than 24 hours': 'sum',
        'More than 72 hours': 'sum',
        'Site Alias': 'nunique'
    }).reset_index()

    pivot = pivot.rename(columns={'Site Alias': 'Total'})

    total_row = pivot[['1 or Less than 1 day', 'More than 24 hours', 'More than 72 hours', 'Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']

    pivot = pd.concat([pivot, total_row], ignore_index=True)

    total_offline_count = int(pivot['Total'].iloc[-1])

    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']

    return pivot, total_offline_count

# Function to calculate time offline smartly (minutes, hours, or days)
def calculate_time_offline(df, current_time):
    df['Last Online Time'] = pd.to_datetime(df['Last Online Time'], format='%Y-%m-%d %H:%M:%S')
    df['Hours Offline'] = (current_time - df['Last Online Time']).dt.total_seconds() / 3600

    def format_offline_duration(hours):
        if hours < 1:
            return f"{int(hours * 60)} minutes"
        elif hours < 24:
            return f"{int(hours)} hours"
        else:
            return f"{int(hours // 24)} days"

    df['Offline Duration'] = df['Hours Offline'].apply(format_offline_duration)

    return df[['Offline Duration', 'Site Alias', 'Cluster', 'Zone', 'Last Online Time']]

# Function to extract the file name's timestamp
def extract_timestamp(file_name):
    match = re.search(r'\((.*?)\)', file_name)
    if match:
        timestamp_str = match.group(1)
        try:
            return pd.to_datetime(timestamp_str.replace('_', ':'), format='%B %dth %Y, %I:%M:%S %p')
        except ValueError:
            return datetime.now()
    return datetime.now()

# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            valid_sheet_name = re.sub(r'[\\/*?:[\]]', '_', sheet_name)[:31]
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    return output.getvalue()

# ====================
# Site Comparison Tool
# ====================

# Function to compare site names and show missing sites
def compare_sites(rms_df, site_access_df):
    # Extract site names from RMS (row 3, column B)
    rms_sites = rms_df.iloc[2:, 1].dropna().str.split('_').str[0].reset_index(drop=True)  # Site names
    rms_alias = rms_df.iloc[2:, 2].dropna().reset_index(drop=True)  # Site Alias from column C
    rms_zone = rms_df.iloc[2:, 3].dropna().reset_index(drop=True)   # Zone from column D
    rms_cluster = rms_df.iloc[2:, 4].dropna().reset_index(drop=True) # Cluster from column E

    # Extract site names from Site Access Portal (Column D, header 'SiteName')
    site_access_sites = site_access_df['SiteName'].str.split('_').str[0]

    # Find missing sites (in RMS but not in Site Access Portal)
    missing_sites = rms_sites[~rms_sites.isin(site_access_sites)].reset_index()

    # Get corresponding Site Alias, Zone, and Cluster for the missing sites
    missing_aliases = rms_alias.loc[missing_sites['index']].reset_index(drop=True)
    missing_zones = rms_zone.loc[missing_sites['index']].reset_index(drop=True)
    missing_clusters = rms_cluster.loc[missing_sites['index']].reset_index(drop=True)

    # Combine the results into a DataFrame for clear output
    result_df = pd.DataFrame({
        'Site Alias': missing_aliases,
        'Zone': missing_zones,
        'Cluster': missing_clusters,
        'Missing Site': missing_sites[0]
    })

    return result_df

# Function to group the data by Cluster and Zone and show the count per Site Alias
def group_cluster_zone_data(df):
    grouped = df.groupby(['Cluster', 'Zone', 'Site Alias']).size().reset_index(name='Count')
    return grouped

# =====================
# Streamlit App Function
# =====================

def main():
    st.set_page_config(page_title="Door Open Alarms & Site Comparison Dashboard", layout="wide")
    st.title("🚪 Door Open Alarms & Site Comparison Dashboard")

    # Sidebar Navigation
    st.sidebar.title("Navigation")
    app_mode = st.sidebar.selectbox("Choose the functionality", ["Alarm Monitoring", "Site Comparison"])

    if app_mode == "Alarm Monitoring":
        # =========================
        # Alarm and Offline Monitoring
        # =========================

        st.header("🔔 Alarm and Offline Data Pivot Table Generator")

        st.sidebar.header("Upload Reports")
        uploaded_alarm_file = st.sidebar.file_uploader("Upload Current Alarms Report", type=["xlsx"])
        uploaded_offline_file = st.sidebar.file_uploader("Upload Offline Report", type=["xlsx"])

        if uploaded_alarm_file and uploaded_offline_file:
            try:
                alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
                offline_df = pd.read_excel(uploaded_offline_file, header=2)

                current_time = extract_timestamp(uploaded_alarm_file.name)
                offline_time = extract_timestamp(uploaded_offline_file.name)

                # Process the Offline Report
                pivot_offline, total_offline_count = create_offline_pivot(offline_df)

                # Display Offline Report
                with st.expander("📄 Offline Report"):
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.markdown(f"**Report Timestamp:** {offline_time.strftime('%Y-%m-%d %H:%M:%S')}")
                    with col2:
                        st.markdown(f"**Total Offline Count:** {total_offline_count}")

                    st.dataframe(pivot_offline.style.highlight_max(axis=0))

                # Calculate time offline smartly using the offline time
                time_offline_df = calculate_time_offline(offline_df, offline_time)

                # Create a summary table based on offline duration
                summary_df = time_offline_df.copy()
                summary_df = summary_df.rename(columns={
                    'Site Alias': 'Site Name',
                    'Last Online Time': 'Last Online'
                })

                # Display the Summary of Offline Sites
                with st.expander("📝 Summary of Offline Sites"):
                    st.dataframe(summary_df)

                # Prepare download for Offline Report
                offline_report_data = {
                    "Offline Summary": pivot_offline,
                    "Offline Details": summary_df
                }
                offline_excel_data = to_excel(offline_report_data)

                st.download_button(
                    label="📥 Download Offline Report",
                    data=offline_excel_data,
                    file_name=f"Offline_Report_{offline_time.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Check for required columns in Alarm Report
                alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name', 'Alarm Time']
                if not all(col in alarm_df.columns for col in alarm_required_columns):
                    st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
                    return

                # Extract Client information
                alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
                alarm_df = alarm_df[~alarm_df['Client'].isnull()]

                # Add the current time to the alarm header
                st.markdown("### 🔔 Current Alarms Report")
                st.markdown(f"**Report Timestamp:** {current_time.strftime('%Y-%m-%d %H:%M:%S')}")

                alarm_names = alarm_df['Alarm Name'].unique()

                # Define the priority order for the alarm names
                priority_order = [
                    'Mains Fail',
                    'Battery Low',
                    'DCDB-01 Primary Disconnect',
                    'PG Run',
                    'MDB Fault',
                    'Door Open'
                ]

                # Separate prioritized alarms from the rest
                prioritized_alarms = [name for name in priority_order if name in alarm_names]
                non_prioritized_alarms = [name for name in alarm_names if name not in priority_order]

                # Combine both lists to maintain the desired order
                ordered_alarm_names = prioritized_alarms + non_prioritized_alarms

                # Create a dictionary to store all pivot tables for current alarms
                alarm_data = {}

                # Add a time filter for the "DCDB-01 Primary Disconnect" alarm
                if 'DCDB-01 Primary Disconnect' in ordered_alarm_names:
                    st.sidebar.subheader("Filter DCDB-01 Primary Disconnect")
                    dcdb_start_date = st.sidebar.date_input("Start Date", value=datetime.now())
                    dcdb_end_date = st.sidebar.date_input("End Date", value=datetime.now())

                for alarm_name in ordered_alarm_names:
                    if alarm_name == 'DCDB-01 Primary Disconnect':
                        # Filter the DataFrame based on the selected date range
                        filtered_data = alarm_df[
                            (alarm_df['Alarm Name'] == alarm_name) &
                            (pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p').dt.date >= dcdb_start_date) &
                            (pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p').dt.date <= dcdb_end_date)
                        ]
                        if not filtered_data.empty:
                            pivot, total_count = create_pivot_table(filtered_data, alarm_name)
                            alarm_data[alarm_name] = (pivot, total_count)
                    else:
                        pivot, total_count = create_pivot_table(alarm_df, alarm_name)
                        alarm_data[alarm_name] = (pivot, total_count)

                # Display each pivot table for the current alarms
                for alarm_name, (pivot, total_count) in alarm_data.items():
                    with st.expander(f"### {alarm_name}"):
                        st.markdown(f"**Alarm Count:** {total_count}")
                        st.dataframe(pivot.style.highlight_max(axis=0))

                # Prepare download for Current Alarms Report only if there is data
                if alarm_data:
                    current_alarm_excel_data = to_excel({alarm_name: data[0] for alarm_name, data in alarm_data.items()})
                    st.download_button(
                        label="📥 Download Current Alarms Report",
                        data=current_alarm_excel_data,
                        file_name=f"Current_Alarms_Report_{current_time.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No current alarm data available for export.")
            except Exception as e:
                st.error(f"An error occurred while processing the files: {e}")
        else:
            st.info("Please upload both the Current Alarms Report and the Offline Report to proceed.")

    elif app_mode == "Site Comparison":
        # ====================
        # Site Comparison Tool
        # ====================

        st.header("📋 Site Comparison Tool")

        st.sidebar.header("Upload Files for Site Comparison")
        uploaded_rms_file = st.sidebar.file_uploader("Upload RMS Excel File", type=["xlsx"])
        uploaded_site_access_file = st.sidebar.file_uploader("Upload Site Access Portal Excel File", type=["xlsx"])

        if uploaded_rms_file and uploaded_site_access_file:
            try:
                rms_df = pd.read_excel(uploaded_rms_file, header=2)  # Header starts from row 3
                site_access_df = pd.read_excel(uploaded_site_access_file)

                # Compare and find missing sites
                missing_sites_df = compare_sites(rms_df, site_access_df)

                if missing_sites_df.empty:
                    st.success("No missing sites found. All sites in RMS are present in the Site Access Portal.")
                else:
                    st.warning(f"Found {len(missing_sites_df)} missing sites.")

                    st.subheader("🔍 Missing Sites Details")
                    st.dataframe(missing_sites_df)

                    # Group the data by Cluster and Zone
                    grouped_df = group_cluster_zone_data(missing_sites_df)

                    st.subheader("📊 Cluster-wise Grouped Data")
                    st.dataframe(grouped_df)

                    # Download options
                    comparison_report = {
                        "Missing Sites": missing_sites_df,
                        "Cluster-wise Grouped Data": grouped_df
                    }
                    comparison_excel_data = to_excel(comparison_report)

                    st.download_button(
                        label="📥 Download Comparison Report",
                        data=comparison_excel_data,
                        file_name=f"Site_Comparison_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"An error occurred while processing the files: {e}")
        else:
            st.info("Please upload both the RMS and Site Access Portal Excel files to proceed.")

if __name__ == "__main__":
    main()
