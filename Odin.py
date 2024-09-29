import pandas as pd
import streamlit as st
import re
from io import BytesIO
from datetime import datetime

# Function to compare site names and show missing sites with Site Alias, Zone, and Cluster
def compare_sites(rms_df, site_access_df):
    # Extract site names from RMS (assuming 'SiteName' is in column B (index 1))
    rms_sites = rms_df.iloc[:, 1].dropna().str.split('_').str[0].reset_index(drop=True)  # Site names
    rms_alias = rms_df.iloc[:, 2].dropna().reset_index(drop=True)  # Site Alias from column C
    rms_zone = rms_df.iloc[:, 3].dropna().reset_index(drop=True)   # Zone from column D
    rms_cluster = rms_df.iloc[:, 4].dropna().reset_index(drop=True) # Cluster from column E

    # Extract site names from Site Access Portal (assuming 'SiteName' is in column D (index 3))
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

# Function to display grouped data in a structured format
def display_cluster_wise_data(df):
    grouped = group_cluster_zone_data(df)
    
    # Iterate through each cluster
    clusters = grouped['Cluster'].unique()
    for cluster in clusters:
        st.markdown(f"### **Cluster: {cluster}**")
        cluster_data = grouped[grouped['Cluster'] == cluster]
        
        # Iterate through each zone within the cluster
        zones = cluster_data['Zone'].unique()
        for zone in zones:
            st.markdown(f"#### **Zone: {zone}**")
            zone_data = cluster_data[grouped['Zone'] == zone]
            
            # Display the data in a table
            st.table(zone_data[['Site Alias', 'Count']])

# Function to convert multiple DataFrames to Excel with separate sheets (Optional)
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in dfs_dict.items():
            # Clean sheet name to avoid issues
            valid_sheet_name = re.sub(r'[\\/*?:[\]]', '_', sheet_name)[:31]
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    return output.getvalue()

# Main function for the Streamlit app
def main():
    st.title("📋 Site Comparison Tool")
    
    st.sidebar.header("Upload Files for Site Comparison")
    uploaded_rms_file = st.sidebar.file_uploader("Upload RMS CSV File", type=["csv"])
    uploaded_site_access_file = st.sidebar.file_uploader("Upload Site Access Portal CSV File", type=["csv"])
    
    if st.sidebar.button("Compare Sites"):
        if uploaded_rms_file is not None and uploaded_site_access_file is not None:
            try:
                # Read the uploaded CSV files
                rms_df = pd.read_csv(uploaded_rms_file)
                site_access_df = pd.read_csv(uploaded_site_access_file)
                
                # Compare and find missing sites
                missing_sites_df = compare_sites(rms_df, site_access_df)
                
                if missing_sites_df.empty:
                    st.success("✅ No missing sites found. All sites in RMS are present in the Site Access Portal.")
                else:
                    st.warning(f"⚠️ Found **{len(missing_sites_df)}** missing sites.")
                    
                    st.subheader("🔍 Missing Sites Details")
                    st.dataframe(missing_sites_df)
                    
                    st.subheader("📊 Cluster-wise Grouped Data")
                    display_cluster_wise_data(missing_sites_df)
                    
                    # Provide a download button for the comparison report
                    comparison_report = {
                        "Missing Sites": missing_sites_df,
                        "Cluster-wise Grouped Data": group_cluster_zone_data(missing_sites_df)
                    }
                    comparison_excel_data = to_excel(comparison_report)
                    
                    st.download_button(
                        label="📥 Download Comparison Report",
                        data=comparison_excel_data,
                        file_name=f"Site_Comparison_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"🔴 An error occurred while processing the files: {e}")
        else:
            st.error("🔴 Please upload both the RMS and Site Access Portal CSV files.")

if __name__ == "__main__":
    main()
