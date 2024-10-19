import streamlit as st
import pandas as pd

def main():
    st.title("Site Access and RMS Comparison Tool")
    
    # Upload Excel files
    site_access_file = st.file_uploader("Upload Site Access Excel File", type=["xlsx"])
    rms_file = st.file_uploader("Upload RMS Excel File", type=["xlsx"])
    
    if site_access_file is not None and rms_file is not None:
        # Read the uploaded Excel files
        site_access_df = pd.read_excel(site_access_file)
        rms_df = pd.read_excel(rms_file)
        
        # Display the uploaded dataframes
        st.subheader("Site Access Data")
        st.dataframe(site_access_df)
        
        st.subheader("RMS Data")
        st.dataframe(rms_df)

        # Check for necessary columns
        required_site_access_columns = ['RequestId', 'SiteName', 'SiteAccessType', 
                                         'StartDate', 'EndDate', 'InTime', 'OutTime', 
                                         'AccessPurpose', 'VendorName', 'POCName']
        required_rms_columns = ['RMS Station', 'Site', 'Site Alias', 'Zone', 
                                'Cluster', 'Alarm Name', 'Tag', 'Tenant', 
                                'Alarm Time', 'Duration', 'Duration Slot (Hours)']

        if all(col in site_access_df.columns for col in required_site_access_columns) and \
           all(col in rms_df.columns for col in required_rms_columns):
            
            # Extract the first part of SiteName before the underscore
            site_access_df['ShortSiteName'] = site_access_df['SiteName'].str.split('_').str[0]
            
            # Find mismatched sites
            mismatches = site_access_df[~site_access_df['ShortSiteName'].isin(rms_df['Site'])]
            
            if not mismatches.empty:
                # Merge to get corresponding Site Alias
                merged_df = pd.merge(mismatches[['ShortSiteName', 'SiteName']], 
                                      rms_df[['Site', 'Site Alias', 'Zone', 'Cluster']], 
                                      left_on='ShortSiteName', right_on='Site', 
                                      how='left')

                # Group by Zone and Cluster
                grouped = merged_df.groupby(['Zone', 'Cluster']).agg({
                    'SiteName': 'unique',
                    'Site Alias': 'unique'
                }).reset_index()

                st.subheader("Mismatched Sites Grouped by Zone and Cluster")
                st.dataframe(grouped)

            else:
                st.success("No mismatches found between Site Access and RMS data.")
        else:
            st.error("One or more required columns are missing in the uploaded Excel files.")

if __name__ == "__main__":
    main()
