import pandas as pd
import streamlit as st

# Function to compare site names and show missing sites with Site Alias, Zone, and Cluster
def compare_sites(rms_df, site_access_df):
    # Extract site names from RMS (row 3, column B)
    rms_sites = rms_df.iloc[2:, 1].dropna().str.split('_').str[0]  # Site names
    rms_alias = rms_df.iloc[2:, 2].dropna()  # Site Alias from column C
    rms_zone = rms_df.iloc[2:, 3].dropna()   # Zone from column D
    rms_cluster = rms_df.iloc[2:, 4].dropna() # Cluster from column E

    # Extract site names from Site Access Portal (Column D, header 'SiteName')
    site_access_sites = site_access_df['SiteName'].str.split('_').str[0]

    # Find missing sites (in RMS but not in Site Access Portal)
    missing_sites = rms_sites[~rms_sites.isin(site_access_sites)]

    # Get corresponding Site Alias, Zone, and Cluster for the missing sites
    missing_aliases = rms_alias[rms_sites.index.isin(missing_sites.index)]
    missing_zones = rms_zone[rms_sites.index.isin(missing_sites.index)]
    missing_clusters = rms_cluster[rms_sites.index.isin(missing_sites.index)]

    # Combine the results into a DataFrame for clear output
    result_df = pd.DataFrame({
        'Site Alias': missing_aliases,
        'Zone': missing_zones,
        'Cluster': missing_clusters,
        'Missing Site': missing_sites
    }).reset_index(drop=True)

    # Return the result DataFrame
    return result_df

# Function to group the data by Cluster and Zone and show the count per Site Alias
def display_cluster_wise_data(df):
    grouped = df.groupby(['Cluster', 'Zone', 'Site Alias']).size().reset_index(name='Count')
    
    # Display grouped data by Cluster first
    clusters = grouped['Cluster'].unique()
    for cluster in clusters:
        st.write(f"### {cluster}\n")
        cluster_data = grouped[grouped['Cluster'] == cluster]
        
        zones = cluster_data['Zone'].unique()
        for zone in zones:
            st.write(f"#### {zone}\n")
            zone_data = cluster_data[cluster_data['Zone'] == zone]
            st.write("Site Name | Count")
            st.write("--- | ---")
            for index, row in zone_data.iterrows():
                st.write(f"{row['Site Alias']} | {row['Count']}")

# Main function for the Streamlit app
def main():
    st.title("Site Comparison Tool")

    # File uploaders for RMS and Site Access Portal files
    rms_file = st.file_uploader("Upload RMS Excel file", type=["xlsx", "xls"])
    site_access_file = st.file_uploader("Upload Site Access Portal Excel file", type=["xlsx", "xls"])

    if rms_file is not None and site_access_file is not None:
        # Read the uploaded files
        rms_df = pd.read_excel(rms_file, header=2)
        site_access_df = pd.read_excel(site_access_file)

        # Compare and find missing sites
        missing_sites_df = compare_sites(rms_df, site_access_df)
        
        # Display the cluster-wise grouped data
        st.write("## Cluster-wise Grouped Data")
        display_cluster_wise_data(missing_sites_df)

if __name__ == "__main__":
    main()
