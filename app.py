import streamlit as st
import pandas as pd
import os
import io

def convert_mib_to_gb(mib_value):
    """Convert MiB to GB"""
    try:
        return round(float(mib_value) / 953.7, 2)
    except (ValueError, TypeError):
        return 0

def convert_mb_to_gb(mb_value):
    """Convert MB to GB"""
    try:
        return round(float(mb_value) / 1024, 2)
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
        
        # Convert memory (MB to GB) and disk (MiB to GB)
        server_list['Memory (GB)'] = server_list['Memory'].apply(convert_mb_to_gb)
        server_list['Provisioned Disk (GB)'] = server_list['Provisioned MB'].apply(convert_mib_to_gb)
        server_list['In Use Disk (GB)'] = server_list['In Use MB'].apply(convert_mib_to_gb)
        
        # Add new columns
        server_list['In Scope for Prod?'] = ''
        server_list['In Scope for DR?'] = ''
        server_list['Notes'] = ''
        
        # Reorder columns
        final_columns = [
            'VM Name',
            'Powerstate',
            'CPUs',
            'Memory (GB)',
            'Provisioned Disk (GB)',
            'In Use Disk (GB)',
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
        try:
            # Process the file
            server_list = process_rvtools_file(uploaded_file)
            
            if server_list is not None:
                # Display the processed data
                st.write("### Processed Server List")
                st.dataframe(server_list)
                
                # Use a generic filename for the download
                output_filename = "rvtools-processed.xlsx"
                
                # Create Excel writer with BytesIO
                excel_file = io.BytesIO()
                output = pd.ExcelWriter(excel_file, engine='openpyxl')
                
                # Write the ServerList tab
                server_list.to_excel(output, sheet_name='ServerList', index=False)
                
                # Get the ServerList worksheet and add filters
                serverlist_ws = output.sheets['ServerList']
                serverlist_ws.auto_filter.ref = serverlist_ws.dimensions
                
                # Create Summary tab
                summary_ws = output.book.create_sheet('Summary')
                
                # Add headers to Summary tab
                summary_headers = [
                    'Category', 'Sub-Category', 'Count', 'Total CPUs', 
                    'Total Memory (GB)', 'Total Provisioned Disk (GB)', 
                    'Total In Use Disk (GB)'
                ]
                for col, header in enumerate(summary_headers, 1):
                    summary_ws.cell(row=1, column=col, value=header)
                
                # Add Powerstate summary
                current_row = 2
                powerstates = server_list['Powerstate'].unique()
                for powerstate in powerstates:
                    # Main category (Powerstate)
                    summary_ws.cell(row=current_row, column=1, value='Powerstate')
                    summary_ws.cell(row=current_row, column=2, value=powerstate)
                    
                    # Count formula
                    summary_ws.cell(row=current_row, column=3, 
                        value=f'=COUNTIFS(ServerList!B:B,"{powerstate}")')
                    
                    # CPU sum formula
                    summary_ws.cell(row=current_row, column=4,
                        value=f'=SUMIFS(ServerList!C:C,ServerList!B:B,"{powerstate}")')
                    
                    # Memory sum formula
                    summary_ws.cell(row=current_row, column=5,
                        value=f'=SUMIFS(ServerList!D:D,ServerList!B:B,"{powerstate}")')
                    
                    # Provisioned Disk sum formula
                    summary_ws.cell(row=current_row, column=6,
                        value=f'=SUMIFS(ServerList!E:E,ServerList!B:B,"{powerstate}")')
                    
                    # In Use Disk sum formula
                    summary_ws.cell(row=current_row, column=7,
                        value=f'=SUMIFS(ServerList!F:F,ServerList!B:B,"{powerstate}")')
                    
                    current_row += 1
                
                # Add Powerstate subtotal
                summary_ws.cell(row=current_row, column=1, value='Powerstate Subtotal')
                summary_ws.cell(row=current_row, column=2, value='')
                summary_ws.cell(row=current_row, column=3, value='=SUM(C2:C' + str(current_row-1) + ')')
                summary_ws.cell(row=current_row, column=4, value='=SUM(D2:D' + str(current_row-1) + ')')
                summary_ws.cell(row=current_row, column=5, value='=SUM(E2:E' + str(current_row-1) + ')')
                summary_ws.cell(row=current_row, column=6, value='=SUM(F2:F' + str(current_row-1) + ')')
                summary_ws.cell(row=current_row, column=7, value='=SUM(G2:G' + str(current_row-1) + ')')
                current_row += 2  # Add extra blank row after subtotal
                
                # Add OS summary
                current_row += 1  # Add a blank row between sections
                summary_ws.cell(row=current_row, column=1, value='Operating System Summary')
                current_row += 1
                
                operating_systems = server_list['OS according to the configuration file'].unique()
                for os_name in operating_systems:
                    # Main category (OS)
                    summary_ws.cell(row=current_row, column=1, value='Operating System')
                    summary_ws.cell(row=current_row, column=2, value=os_name)
                    
                    # Count formula
                    summary_ws.cell(row=current_row, column=3,
                        value=f'=COUNTIFS(ServerList!H:H,"{os_name}")')
                    
                    # CPU sum formula
                    summary_ws.cell(row=current_row, column=4,
                        value=f'=SUMIFS(ServerList!C:C,ServerList!H:H,"{os_name}")')
                    
                    # Memory sum formula
                    summary_ws.cell(row=current_row, column=5,
                        value=f'=SUMIFS(ServerList!D:D,ServerList!H:H,"{os_name}")')
                    
                    # Provisioned Disk sum formula
                    summary_ws.cell(row=current_row, column=6,
                        value=f'=SUMIFS(ServerList!E:E,ServerList!H:H,"{os_name}")')
                    
                    # In Use Disk sum formula
                    summary_ws.cell(row=current_row, column=7,
                        value=f'=SUMIFS(ServerList!F:F,ServerList!H:H,"{os_name}")')
                    
                    current_row += 1
                
                # Add grand totals
                current_row += 1  # Add a blank row
                summary_ws.cell(row=current_row, column=1, value='Grand Total')
                summary_ws.cell(row=current_row, column=3, value='=COUNTA(ServerList!A:A)-1')  # Subtract 1 for header
                summary_ws.cell(row=current_row, column=4, value='=SUM(ServerList!C:C)')
                summary_ws.cell(row=current_row, column=5, value='=SUM(ServerList!D:D)')
                summary_ws.cell(row=current_row, column=6, value='=SUM(ServerList!E:E)')
                summary_ws.cell(row=current_row, column=7, value='=SUM(ServerList!F:F)')
                
                # Add Prod Scope summary
                current_row += 2  # Add blank rows between sections
                summary_ws.cell(row=current_row, column=1, value='Production Scope Summary')
                current_row += 1
                
                # Add "In Scope" row
                summary_ws.cell(row=current_row, column=1, value='Production Scope')
                summary_ws.cell(row=current_row, column=2, value='In Scope')
                
                # Count formula for In Scope (checking for various boolean values)
                count_formula = '=COUNTIFS(ServerList!I:I,"yes")+COUNTIFS(ServerList!I:I,"true")+COUNTIFS(ServerList!I:I,"1")+COUNTIFS(ServerList!I:I,"X")+COUNTIFS(ServerList!I:I,"y")'
                summary_ws.cell(row=current_row, column=3, value=count_formula)
                
                # CPU sum formula for In Scope
                cpu_formula = '=SUMIFS(ServerList!C:C,ServerList!I:I,"yes")+SUMIFS(ServerList!C:C,ServerList!I:I,"true")+SUMIFS(ServerList!C:C,ServerList!I:I,"1")+SUMIFS(ServerList!C:C,ServerList!I:I,"X")+SUMIFS(ServerList!C:C,ServerList!I:I,"y")'
                summary_ws.cell(row=current_row, column=4, value=cpu_formula)
                
                # Memory sum formula for In Scope
                memory_formula = '=SUMIFS(ServerList!D:D,ServerList!I:I,"yes")+SUMIFS(ServerList!D:D,ServerList!I:I,"true")+SUMIFS(ServerList!D:D,ServerList!I:I,"1")+SUMIFS(ServerList!D:D,ServerList!I:I,"X")+SUMIFS(ServerList!D:D,ServerList!I:I,"y")'
                summary_ws.cell(row=current_row, column=5, value=memory_formula)
                
                # Provisioned Disk sum formula for In Scope
                disk_prov_formula = '=SUMIFS(ServerList!E:E,ServerList!I:I,"yes")+SUMIFS(ServerList!E:E,ServerList!I:I,"true")+SUMIFS(ServerList!E:E,ServerList!I:I,"1")+SUMIFS(ServerList!E:E,ServerList!I:I,"X")+SUMIFS(ServerList!E:E,ServerList!I:I,"y")'
                summary_ws.cell(row=current_row, column=6, value=disk_prov_formula)
                
                # In Use Disk sum formula for In Scope
                disk_used_formula = '=SUMIFS(ServerList!F:F,ServerList!I:I,"yes")+SUMIFS(ServerList!F:F,ServerList!I:I,"true")+SUMIFS(ServerList!F:F,ServerList!I:I,"1")+SUMIFS(ServerList!F:F,ServerList!I:I,"X")+SUMIFS(ServerList!F:F,ServerList!I:I,"y")'
                summary_ws.cell(row=current_row, column=7, value=disk_used_formula)
                
                current_row += 1
                
                # Add "Not In Scope" row for Production
                summary_ws.cell(row=current_row, column=1, value='Production Scope')
                summary_ws.cell(row=current_row, column=2, value='Not In Scope')
                
                # Count formula for Not In Scope (everything except yes/true/1/X)
                count_formula = '=COUNTIFS(ServerList!I:I,"<>yes",ServerList!I:I,"<>true",ServerList!I:I,"<>1",ServerList!I:I,"<>X",ServerList!I:I,"<>y")'
                summary_ws.cell(row=current_row, column=3, value=count_formula)
                
                # CPU sum formula for Not In Scope
                cpu_formula = '=SUMIFS(ServerList!C:C,ServerList!I:I,"<>yes",ServerList!I:I,"<>true",ServerList!I:I,"<>1",ServerList!I:I,"<>X",ServerList!I:I,"<>y")'
                summary_ws.cell(row=current_row, column=4, value=cpu_formula)
                
                # Memory sum formula for Not In Scope
                memory_formula = '=SUMIFS(ServerList!D:D,ServerList!I:I,"<>yes",ServerList!I:I,"<>true",ServerList!I:I,"<>1",ServerList!I:I,"<>X",ServerList!I:I,"<>y")'
                summary_ws.cell(row=current_row, column=5, value=memory_formula)
                
                # Provisioned Disk sum formula for Not In Scope
                disk_prov_formula = '=SUMIFS(ServerList!E:E,ServerList!I:I,"<>yes",ServerList!I:I,"<>true",ServerList!I:I,"<>1",ServerList!I:I,"<>X",ServerList!I:I,"<>y")'
                summary_ws.cell(row=current_row, column=6, value=disk_prov_formula)
                
                # In Use Disk sum formula for Not In Scope
                disk_used_formula = '=SUMIFS(ServerList!F:F,ServerList!I:I,"<>yes",ServerList!I:I,"<>true",ServerList!I:I,"<>1",ServerList!I:I,"<>X",ServerList!I:I,"<>y")'
                summary_ws.cell(row=current_row, column=7, value=disk_used_formula)
                
                current_row += 1
                
                # Add Production Scope subtotal
                prod_start_row = current_row - 2  # Row where Production section started
                summary_ws.cell(row=current_row, column=1, value='Production Scope Subtotal')
                summary_ws.cell(row=current_row, column=2, value='')
                summary_ws.cell(row=current_row, column=3, value=f'=SUM(C{prod_start_row}:C{current_row-1})')
                summary_ws.cell(row=current_row, column=4, value=f'=SUM(D{prod_start_row}:D{current_row-1})')
                summary_ws.cell(row=current_row, column=5, value=f'=SUM(E{prod_start_row}:E{current_row-1})')
                summary_ws.cell(row=current_row, column=6, value=f'=SUM(F{prod_start_row}:F{current_row-1})')
                summary_ws.cell(row=current_row, column=7, value=f'=SUM(G{prod_start_row}:G{current_row-1})')
                current_row += 2  # Add extra blank row after subtotal
                
                # Add DR Scope summary
                current_row += 1  # Add blank row between sections
                summary_ws.cell(row=current_row, column=1, value='DR Scope Summary')
                current_row += 1
                
                # Add "In Scope" row
                summary_ws.cell(row=current_row, column=1, value='DR Scope')
                summary_ws.cell(row=current_row, column=2, value='In Scope')
                
                # Count formula for In Scope (checking for various boolean values)
                count_formula = '=COUNTIFS(ServerList!J:J,"yes")+COUNTIFS(ServerList!J:J,"true")+COUNTIFS(ServerList!J:J,"1")+COUNTIFS(ServerList!J:J,"X")+COUNTIFS(ServerList!J:J,"y")'
                summary_ws.cell(row=current_row, column=3, value=count_formula)
                
                # CPU sum formula for In Scope
                cpu_formula = '=SUMIFS(ServerList!C:C,ServerList!J:J,"yes")+SUMIFS(ServerList!C:C,ServerList!J:J,"true")+SUMIFS(ServerList!C:C,ServerList!J:J,"1")+SUMIFS(ServerList!C:C,ServerList!J:J,"X")+SUMIFS(ServerList!C:C,ServerList!J:J,"y")'
                summary_ws.cell(row=current_row, column=4, value=cpu_formula)
                
                # Memory sum formula for In Scope
                memory_formula = '=SUMIFS(ServerList!D:D,ServerList!J:J,"yes")+SUMIFS(ServerList!D:D,ServerList!J:J,"true")+SUMIFS(ServerList!D:D,ServerList!J:J,"1")+SUMIFS(ServerList!D:D,ServerList!J:J,"X")+SUMIFS(ServerList!D:D,ServerList!J:J,"y")'
                summary_ws.cell(row=current_row, column=5, value=memory_formula)
                
                # Provisioned Disk sum formula for In Scope
                disk_prov_formula = '=SUMIFS(ServerList!E:E,ServerList!J:J,"yes")+SUMIFS(ServerList!E:E,ServerList!J:J,"true")+SUMIFS(ServerList!E:E,ServerList!J:J,"1")+SUMIFS(ServerList!E:E,ServerList!J:J,"X")+SUMIFS(ServerList!E:E,ServerList!J:J,"y")'
                summary_ws.cell(row=current_row, column=6, value=disk_prov_formula)
                
                # In Use Disk sum formula for In Scope
                disk_used_formula = '=SUMIFS(ServerList!F:F,ServerList!J:J,"yes")+SUMIFS(ServerList!F:F,ServerList!J:J,"true")+SUMIFS(ServerList!F:F,ServerList!J:J,"1")+SUMIFS(ServerList!F:F,ServerList!J:J,"X")+SUMIFS(ServerList!F:F,ServerList!J:J,"y")'
                summary_ws.cell(row=current_row, column=7, value=disk_used_formula)
                
                current_row += 1
                
                # Add "Not In Scope" row for DR
                summary_ws.cell(row=current_row, column=1, value='DR Scope')
                summary_ws.cell(row=current_row, column=2, value='Not In Scope')
                
                # Count formula for Not In Scope (everything except yes/true/1/X)
                count_formula = '=COUNTIFS(ServerList!J:J,"<>yes",ServerList!J:J,"<>true",ServerList!J:J,"<>1",ServerList!J:J,"<>X",ServerList!J:J,"<>y")'
                summary_ws.cell(row=current_row, column=3, value=count_formula)
                
                # CPU sum formula for Not In Scope
                cpu_formula = '=SUMIFS(ServerList!C:C,ServerList!J:J,"<>yes",ServerList!J:J,"<>true",ServerList!J:J,"<>1",ServerList!J:J,"<>X",ServerList!J:J,"<>y")'
                summary_ws.cell(row=current_row, column=4, value=cpu_formula)
                
                # Memory sum formula for Not In Scope
                memory_formula = '=SUMIFS(ServerList!D:D,ServerList!J:J,"<>yes",ServerList!J:J,"<>true",ServerList!J:J,"<>1",ServerList!J:J,"<>X",ServerList!J:J,"<>y")'
                summary_ws.cell(row=current_row, column=5, value=memory_formula)
                
                # Provisioned Disk sum formula for Not In Scope
                disk_prov_formula = '=SUMIFS(ServerList!E:E,ServerList!J:J,"<>yes",ServerList!J:J,"<>true",ServerList!J:J,"<>1",ServerList!J:J,"<>X",ServerList!J:J,"<>y")'
                summary_ws.cell(row=current_row, column=6, value=disk_prov_formula)
                
                # In Use Disk sum formula for Not In Scope
                disk_used_formula = '=SUMIFS(ServerList!F:F,ServerList!J:J,"<>yes",ServerList!J:J,"<>true",ServerList!J:J,"<>1",ServerList!J:J,"<>X",ServerList!J:J,"<>y")'
                summary_ws.cell(row=current_row, column=7, value=disk_used_formula)
                
                current_row += 1
                
                # Add DR Scope subtotal
                dr_start_row = current_row - 2  # Row where DR section started
                summary_ws.cell(row=current_row, column=1, value='DR Scope Subtotal')
                summary_ws.cell(row=current_row, column=2, value='')
                summary_ws.cell(row=current_row, column=3, value=f'=SUM(C{dr_start_row}:C{current_row-1})')
                summary_ws.cell(row=current_row, column=4, value=f'=SUM(D{dr_start_row}:D{current_row-1})')
                summary_ws.cell(row=current_row, column=5, value=f'=SUM(E{dr_start_row}:E{current_row-1})')
                summary_ws.cell(row=current_row, column=6, value=f'=SUM(F{dr_start_row}:F{current_row-1})')
                summary_ws.cell(row=current_row, column=7, value=f'=SUM(G{dr_start_row}:G{current_row-1})')
                current_row += 2  # Add extra blank row after subtotal
                
                # Format the Summary tab
                for col in range(1, len(summary_headers) + 1):
                    summary_ws.column_dimensions[chr(64 + col)].width = 20
                
                # Save to BytesIO instead of a file
                output.save(excel_file)
                excel_file.seek(0)
                
                # Create download button with the in-memory file
                st.download_button(
                    label="Download processed Excel file",
                    data=excel_file,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Close the BytesIO object
                excel_file.close()
                
                # Close the Excel writer
                output.close()
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

if __name__ == "__main__":
    main() 