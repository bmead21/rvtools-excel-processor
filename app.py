import streamlit as st
import pandas as pd
import os

def convert_mib_to_gib(mib_value):
    """Convert MiB to GiB"""
    try:
        return round(float(mib_value) / 1024, 2)
    except (ValueError, TypeError):
        return 0

def find_column(df, possible_names):
    """Find a column in the dataframe using possible name variations"""
    for name in possible_names:
        if name in df.columns:
            return name
    return None

def process_rvtools_file(uploaded_file):
    """Process the RVtools Excel file and create a new ServerList tab"""
    try:
        # Read the vInfo tab
        df = pd.read_excel(uploaded_file, sheet_name='vInfo')
        
        # Define possible column name variations
        column_mappings = {
            'VM Name': ['VM Name', 'Name', 'Virtual Machine Name', 'VMName', 'VM'],
            'Powerstate': ['Powerstate', 'Power State', 'Power', 'State'],
            'CPUs': ['CPUs', 'CPU', 'Num CPU', 'vCPUs'],
            'Memory': ['Memory', 'Memory MB', 'Memory (MB)', 'RAM'],
            'Provisioned MB': ['Provisioned MB', 'Provisioned MiB', 'Provisioned', 'Provisioned Storage', 'Provisioned Space'],
            'In Use MB': ['In Use MB', 'In Use MiB', 'Used Space', 'Used Storage', 'In Use Space'],
            'Cluster': ['Cluster', 'vSphere Cluster', 'ESX Cluster'],
            'OS according to the configuration file': ['OS according to the configuration file', 'OS According to the configuration file', 'Guest OS', 'Operating System', 'OS']
        }
        
        # Create a dictionary to store the actual column names found
        found_columns = {}
        for target_col, possible_names in column_mappings.items():
            found_col = find_column(df, possible_names)
            if found_col is None:
                st.error(f"Could not find column for {target_col}. Available columns are: {', '.join(df.columns)}")
                return None
            found_columns[target_col] = found_col
        
        # Select and rename required columns
        server_list = df[[
            found_columns['VM Name'],
            found_columns['Powerstate'],
            found_columns['CPUs'],
            found_columns['Memory'],
            found_columns['Provisioned MB'],
            found_columns['In Use MB'],
            found_columns['Cluster'],
            found_columns['OS according to the configuration file']
        ]].copy()
        
        # Rename columns to standard names
        server_list.columns = [
            'VM Name',
            'Powerstate',
            'CPUs',
            'Memory',
            'Provisioned MB',
            'In Use MB',
            'Cluster',
            'OS according to the configuration file'
        ]
        
        # Convert MiB to GiB for memory and disk columns
        server_list['Memory (GiB)'] = server_list['Memory'].apply(convert_mib_to_gib)
        server_list['Provisioned Disk (GiB)'] = server_list['Provisioned MB'].apply(convert_mib_to_gib)
        server_list['In Use Disk (GiB)'] = server_list['In Use MB'].apply(convert_mib_to_gib)
        
        # Add new columns
        server_list['In Scope for Prod?'] = ''
        server_list['In Scope for DR?'] = ''
        server_list['Notes'] = ''
        
        # Reorder columns
        final_columns = [
            'VM Name',
            'Powerstate',
            'CPUs',
            'Memory (GiB)',
            'Provisioned Disk (GiB)',
            'In Use Disk (GiB)',
            'Cluster',
            'OS according to the configuration file',
            'In Scope for Prod?',
            'In Scope for DR?',
            'Notes'
        ]
        
        server_list = server_list[final_columns]
        
        return server_list
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.error("Available columns in the file: " + ", ".join(df.columns))
        return None

def main():
    st.title("RVtools Excel Processor")
    st.write("Upload an RVtools Excel file to create a ServerList tab")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Process the file
        server_list = process_rvtools_file(uploaded_file)
        
        if server_list is not None:
            # Display the processed data
            st.write("### Processed Server List")
            st.dataframe(server_list)
            
            # Add download button
            output = pd.ExcelWriter('processed_rvtools.xlsx', engine='openpyxl')
            server_list.to_excel(output, sheet_name='ServerList', index=False)
            output.close()
            
            with open('processed_rvtools.xlsx', 'rb') as f:
                st.download_button(
                    label="Download processed Excel file",
                    data=f,
                    file_name="processed_rvtools.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main() 