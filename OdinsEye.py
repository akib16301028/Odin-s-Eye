import os
import streamlit as st
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

# Set the title of the application
st.title("Tenant-Wise Data Processing Application")

# Sidebar for uploading files
st.sidebar.header("Upload Required Excel Files")

# Function to standardize tenant names
def standardize_tenant(tenant_name):
    tenant_mapping = {
        "BANJO": "Banjo",
        "BL": "Banglalink",
        "GP": "Grameenphone",
        "ROBI": "Robi",
    }
    return tenant_mapping.get(tenant_name, tenant_name)

# Function to extract tenant from Site Alias
def extract_tenant(site_alias):
    if isinstance(site_alias, str):
        brackets = site_alias.split("(")
        tenants = [part.split(")")[0].strip() for part in brackets if ")" in part]
        for tenant in tenants:
            if "BANJO" in tenant:
                return "Banjo"
        return tenants[0] if tenants else "Unknown"
    return "Unknown"

# Function to convert elapsed time to decimal hours
def convert_to_decimal_hours(elapsed_time):
    if pd.notnull(elapsed_time):
        total_seconds = elapsed_time.total_seconds()
        decimal_hours = total_seconds / 3600
        return Decimal(decimal_hours).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP)
    return Decimal(0.0)

# Step 1: Upload RMS Site List
rms_site_file = st.sidebar.file_uploader("1. RMS Site List", type=["xlsx", "xls"])
if rms_site_file:
    st.success("RMS Site List uploaded successfully!")
    try:
        df_rms_site = pd.read_excel(rms_site_file, skiprows=2)
        df_rms_filtered = df_rms_site[~df_rms_site["Site"].str.startswith("L", na=False)]
        df_rms_filtered["Tenant"] = df_rms_filtered["Site Alias"].apply(extract_tenant)
        df_rms_filtered["Tenant"] = df_rms_filtered["Tenant"].apply(standardize_tenant)

        tenant_zone_rms = {}
        for tenant in df_rms_filtered["Tenant"].unique():
            tenant_df = df_rms_filtered[df_rms_filtered["Tenant"] == tenant]
            grouped_df = tenant_df.groupby(["Cluster", "Zone"]).size().reset_index(name="Total Site Count")
            grouped_df = grouped_df.sort_values(by=["Cluster", "Zone"])
            tenant_zone_rms[tenant] = grouped_df

    except Exception as e:
        st.error(f"Error processing RMS Site List: {e}")

# Fetch MTA Site List
mta_site_path = os.path.join(os.path.dirname(__file__), "MTA Site List.xlsx")
if os.path.exists(mta_site_path):
    try:
        df_mta_site = pd.read_excel(mta_site_path, skiprows=0)
    except Exception as e:
        st.error(f"Error reading MTA Site List file: {e}")
else:
    st.error("MTA Site List file not found in the repository.")

# Additional table for MTA Sites
if rms_site_file and os.path.exists(mta_site_path):
    try:
        # Group clusters and zones for MTA Sites
        mta_grouped = df_mta_site.groupby(["Cluster", "Zone"])["Site Alias"].count().reset_index()
        mta_grouped.rename(columns={"Site Alias": "Total Site Count"}, inplace=True)

        # Match Yesterday Alarm History and calculate metrics
        alarm_history_file = st.sidebar.file_uploader("2. Yesterday Alarm History", type=["xlsx", "xls"])
        if alarm_history_file:
            df_alarm_history = pd.read_excel(alarm_history_file, skiprows=2)
            matched_alarm = df_alarm_history[df_alarm_history["Site Alias"].isin(df_mta_site["Site Alias"])]
            grouped_alarm = matched_alarm.groupby(["Cluster", "Zone"]).size().reset_index(name="Total Affected Site")
            matched_alarm["Elapsed Time"] = pd.to_timedelta(matched_alarm["Elapsed Time"], errors="coerce")
            elapsed_time_sum = matched_alarm.groupby(["Cluster", "Zone"])["Elapsed Time"].sum().reset_index()
            elapsed_time_sum["Elapsed Time (Decimal)"] = elapsed_time_sum["Elapsed Time"].apply(convert_to_decimal_hours)

            # Match Grid Data
            grid_data_file = st.sidebar.file_uploader("3. Grid Data", type=["xlsx", "xls"])
            if grid_data_file:
                df_grid_data = pd.read_excel(grid_data_file, sheet_name="Site Wise Summary", skiprows=2)
                matched_grid = df_grid_data[df_grid_data["Site Alias"].isin(df_mta_site["Site Alias"])]
                grid_availability = matched_grid.groupby(["Cluster", "Zone"])["AC Availability (%)"].mean().reset_index()

                # Match Total Elapsed
                total_elapse_file = st.sidebar.file_uploader("4. Total Elapse Till Date", type=["xlsx", "xls", "csv"])
                if total_elapse_file:
                    if total_elapse_file.name.endswith(".csv"):
                        df_total_elapse = pd.read_csv(total_elapse_file)
                    else:
                        df_total_elapse = pd.read_excel(total_elapse_file, skiprows=0)
                    matched_elapsed = df_total_elapse[df_total_elapse["Site Alias"].isin(df_mta_site["Site Alias"])]
                    total_redeemed = matched_elapsed.groupby(["Cluster", "Zone"])["Elapsed Time"].sum().reset_index()
                    total_redeemed["Total Redeemed Hour"] = total_redeemed["Elapsed Time"].apply(convert_to_decimal_hours)

                    # Combine everything for MTA table
                    mta_final = pd.merge(mta_grouped, grouped_alarm, on=["Cluster", "Zone"], how="left")
                    mta_final = pd.merge(mta_final, elapsed_time_sum[["Cluster", "Zone", "Elapsed Time (Decimal)"]], on=["Cluster", "Zone"], how="left")
                    mta_final = pd.merge(mta_final, grid_availability, on=["Cluster", "Zone"], how="left")
                    mta_final = pd.merge(mta_final, total_redeemed[["Cluster", "Zone", "Total Redeemed Hour"]], on=["Cluster", "Zone"], how="left")

                    mta_final["Total Allowable Limit (Hr)"] = mta_final["Total Site Count"] * 24 * 30 * (1 - 0.9985)
                    mta_final["Remaining Hour"] = mta_final["Total Allowable Limit (Hr)"] - mta_final["Total Redeemed Hour"]

                    st.subheader("MTA Sites - Final Merged Table")
                    st.dataframe(
                        mta_final[
                            [
                                "Cluster",
                                "Zone",
                                "Total Site Count",
                                "Total Affected Site",
                                "Elapsed Time (Decimal)",
                                "AC Availability (%)",
                                "Total Redeemed Hour",
                                "Total Allowable Limit (Hr)",
                                "Remaining Hour"
                            ]
                        ]
                    )
    except Exception as e:
        st.error(f"Error processing MTA Sites table: {e}")
