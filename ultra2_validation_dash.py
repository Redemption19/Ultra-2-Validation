import streamlit as st
import pandas as pd
import os
from pathlib import Path
import subprocess
from openpyxl import load_workbook
from openpyxl.styles import Font

class DashboardUtils:
    @staticmethod
    def select_folder(key_suffix="default"):
        """Provides folder path input functionality with persistent state"""
        folder_path = st.text_input(
            "Enter folder path:",
            value=st.session_state.get('persistent_folder_path', ''),  # Get saved path
            key=f"folder_input_{key_suffix}"
        )
        
        if folder_path and os.path.exists(folder_path):
            # Save the path to session state for persistence
            st.session_state.folder_path = folder_path
            st.session_state.persistent_folder_path = folder_path  # Save for persistence
            return folder_path
        return None

class Dashboard:
    def __init__(self):
        self.setup_page_config()
        self.setup_session_state()
        self.utils = DashboardUtils()
        
    def setup_page_config(self):
        """Configure the Streamlit page settings"""
        st.set_page_config(
            page_title="Excel Processing Dashboard",
            page_icon="📊",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        self.apply_custom_css()

    def setup_session_state(self):
        """Initialize session state variables"""
        if 'persistent_folder_path' not in st.session_state:
            st.session_state.persistent_folder_path = ''
        if 'folder_path' not in st.session_state:
            st.session_state.folder_path = st.session_state.get('persistent_folder_path', '')

    def apply_custom_css(self):
        """Apply custom CSS styling"""
        st.markdown("""
            <style>
            .stButton > button {
                width: 100%;
                background-color: #4CAF50;
                color: white;
                padding: 10px 24px;
                border-radius: 5px;
                border: none;
                cursor: pointer;
            }
            .stButton > button:hover {
                background-color: #45a049;
            }
            .status-box {
                padding: 1rem;
                border-radius: 5px;
                margin: 1rem 0;
            }
            </style>
        """, unsafe_allow_html=True)

    def create_sidebar(self):
        """Create and return the sidebar navigation"""
        with st.sidebar:
            st.title("📊 Schedule Processing Tools ULTRA 2")
            st.markdown("---")
            
            # Folder selection section with persistent path
            st.subheader("📁 Folder Selection")
            current_path = st.session_state.get('persistent_folder_path', '')
            if current_path:
                st.success(f"Current folder: {Path(current_path).name}")
            
            folder_path = self.utils.select_folder("sidebar")
            
            if folder_path:
                st.success(f"Selected: {Path(folder_path).name}")
            elif not current_path:
                st.warning("Please enter a valid folder path")
            
            st.markdown("---")
            
            # Navigation section
            st.subheader("🔍 Functions")
            selected = st.radio(
                "Choose a function:",
                [
                    "Vlookup",
                    "Multiple Ssnit",
                    "Multiple Account Names",
                    "Find My Schedule",
                    "Validation",
                    "Append Total"
                ]
            )
            
            st.markdown("---")
            st.markdown("### ℹ️ About")
            st.info("Process and validate schedule files for pension management.")
            
            return selected

    def get_duplicates(self):
        """Handle duplicate detection functionality"""
        st.header("🔍 Get Duplicates")
        
        folder_path = st.session_state.persistent_folder_path
        if not folder_path:
            st.warning("Please select a folder first!")
            return
            
        try:
            company_name = Path(folder_path).name
            vlookup_path = Path(folder_path) / f"vlookup_{company_name}.xlsx"
            
            if not vlookup_path.exists():
                st.error(f"VLOOKUP file not found: {vlookup_path}")
                return
                
            with st.spinner("Processing duplicates..."):
                df = pd.read_excel(vlookup_path)
                df['FullName'] = df[['Surname', 'First_Name', 'Other_Names']].fillna('').astype(str).agg(' '.join, axis=1)
                duplicates = df[df.duplicated(subset='FullName', keep=False)]
                
                if duplicates.empty:
                    st.success("No duplicates found!")
                else:
                    st.warning(f"Found {len(duplicates)} duplicate entries")
                    st.dataframe(duplicates[['Ssnit', 'FullName']])
                    
                    if st.button("Export Results"):
                        export_path = Path(folder_path) / "duplicates.xlsx"
                        duplicates.to_excel(export_path, index=False)
                        st.success(f"Results exported to: {export_path}")
                        
        except Exception as e:
            st.error(f"Error processing duplicates: {str(e)}")

    def ssnit_search(self):
        """Handle SSNIT search functionality"""
        st.header("🔎 SSNIT Search")
        
        folder_path = st.session_state.persistent_folder_path
        if not folder_path:
            st.warning("Please select a folder first!")
            return
        
        ssnit_number = st.text_input("Enter SSNIT Number:")
        if st.button("Search") and ssnit_number:
            try:
                results = []
                files = list(Path(folder_path).glob("*.xlsx"))
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, file_path in enumerate(files):
                    if file_path.name.startswith(("vlookup_", "duplicate_ssnit_")):
                        continue
                        
                    status_text.text(f"Searching in {file_path.name}...")
                    progress_bar.progress((idx + 1) / len(files))
                    
                    df = pd.read_excel(file_path)
                    if 'ssnit' in df.columns:
                        df['ssnit'] = df['ssnit'].astype(str).str.strip()
                        match = df[df['ssnit'] == ssnit_number]
                        if not match.empty:
                            results.append({
                                'File': file_path.name,
                                'Names': ', '.join(match['name'].tolist())
                            })
                
                progress_bar.empty()
                status_text.empty()
                
                if results:
                    st.success("Matches found!")
                    st.table(pd.DataFrame(results))
                else:
                    st.info("No matches found.")
                    
            except Exception as e:
                st.error(f"Error during search: {str(e)}")

    def vlookup(self):
        """Handle VLOOKUP functionality"""
        st.header("📑 VLOOKUP Generator")
        
        # File upload for almighty report
        st.subheader("1. Upload Almighty Report")
        almighty_file = st.file_uploader(
            "Upload almighty report Excel file",
            type=['xlsx'],
            key="almighty_upload"
        )
        
        # Required columns definition
        required_columns = [
            'Client Account Number', 'Surname', 'First Name', 
            'Other Names', 'Ssnit', 'Accountno', 'Employer Name', 
        ]
        
        folder_path = st.session_state.persistent_folder_path
        if almighty_file is not None:
            try:
                # Read the almighty report
                with st.spinner("Reading almighty report..."):
                    df = pd.read_excel(almighty_file)
                
                # Validate required columns
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    st.error(f"Missing columns in almighty file: {', '.join(missing_columns)}")
                    return
                
                # Folder selection with unique key
                st.subheader("2. Select Company Folder")
                folder_path = self.utils.select_folder("vlookup")
                
                if folder_path:
                    company_name = Path(folder_path).name
                    
                    # Process the data
                    with st.spinner("Processing data..."):
                        # Filter by company name
                        company_df = df[df['Employer Name'] == company_name]
                        
                        if company_df.empty:
                            st.error(f"No data found for company: {company_name}")
                            return
                        
                        # Clean SSNIT numbers
                        company_df.loc[:, 'Ssnit'] = company_df['Ssnit'].str.strip()
                        
                        # Select and rename columns
                        columns_to_keep = [
                            'Client Account Number', 'Surname', 'First Name',
                            'Other Names', 'Ssnit', 'Accountno'
                        ]
                        renamed_columns = [
                            'Accountno', 'Surname', 'First_Name',
                            'Other_Names', 'Ssnit', 'Accountno2'
                        ]
                        
                        selected_columns_df = company_df[columns_to_keep].rename(
                            columns=dict(zip(columns_to_keep, renamed_columns))
                        )
                        
                        # Preview the data
                        st.subheader("3. Preview Generated Data")
                        st.dataframe(selected_columns_df.head())
                        
                        # Export option
                        if st.button("Generate VLOOKUP File", key="generate_vlookup"):
                            output_path = Path(folder_path) / f'vlookup_{company_name}.xlsx'
                            selected_columns_df.to_excel(output_path, index=False)
                            st.success(f"VLOOKUP file has been created: {output_path}")
                            
                            # Display summary
                            st.info(f"""
                                Summary:
                                - Total records: {len(selected_columns_df)}
                                - Company: {company_name}
                                - Output location: {output_path}
                            """)
                            
            except Exception as e:
                st.error(f"Error processing VLOOKUP: {str(e)}")
                
    def run(self):
        """Main method to run the dashboard"""
        selected = self.create_sidebar()
        
        # Use the persistent folder path for all functions
        if not st.session_state.get('persistent_folder_path'):
            st.warning("Please select a folder first!")
            return
            
        if selected == "Vlookup":
            self.vlookup()
        elif selected == "Multiple Ssnit":
            self.multiple_ssnit()
        elif selected == "Multiple Account Names":
            self.multiple_account_names()
        elif selected == "Find My Schedule":
            self.find_my_schedule()
        elif selected == "Validation":
            self.validation()
        elif selected == "Append Total":
            self.append_total()

    def multiple_ssnit(self):
        """Handle Multiple SSNIT functionality"""
        st.header("🔍 Check Multiple SSNIT Numbers")
        
        # Get main company folder path
        main_folder_path = self.utils.select_folder("multiple_ssnit")
        if not main_folder_path:
            st.warning("Please select a folder first!")
            return
        
        company_name = Path(main_folder_path).name
        st.success(f"Selected company folder: {company_name}")
        
        # Check for VLOOKUP file
        vlookup_path = Path(main_folder_path) / f"vlookup_{company_name}.xlsx"
        if not vlookup_path.exists():
            st.error("VLOOKUP file not found! Please run the VLOOKUP process first.")
            return
        
        try:
            # Read VLOOKUP file and process names
            company_df = pd.read_excel(vlookup_path)
            company_df['FullName'] = company_df[['Surname', 'First_Name', 'Other_Names']].fillna('').astype(str).agg(' '.join, axis=1)
            company_df['SortedFullName'] = company_df['FullName'].apply(lambda x: ' '.join(sorted(x.split())))
            company_df.sort_values(by='SortedFullName', inplace=True)
            
            # Find duplicates
            columns_to_keep = ['Ssnit', 'SortedFullName', 'Surname', 'First_Name', 'Other_Names', 'Accountno', 'Accountno2']
            duplicates = company_df[company_df.duplicated(subset='SortedFullName', keep=False)].loc[:, columns_to_keep]
            
            if not duplicates.empty:
                st.warning(f"Found {len(duplicates) // 2} duplicate names!")
                
                # Create dictionary of duplicates
                duplicate_dict = {}
                for idx, row in duplicates.iterrows():
                    name = row['SortedFullName']
                    ssnit = row['Ssnit']
                    if name not in duplicate_dict:
                        duplicate_dict[name] = set()
                    duplicate_dict[name].add(ssnit)
                
                # Display duplicates in expandable sections
                for name, ssnit_set in duplicate_dict.items():
                    with st.expander(f"🔍 {name} ({len(ssnit_set)} SSNITs)"):
                        # Show the duplicate entries
                        st.dataframe(
                            duplicates[duplicates['SortedFullName'] == name],
                            column_config={
                                "Ssnit": "SSNIT Number",
                                "Surname": "Surname",
                                "First_Name": "First Name",
                                "Other_Names": "Other Names",
                                "Accountno": "Account Number",
                                "Accountno2": "Account Number 2"
                            },
                            hide_index=True
                        )
                        
                        # Check schedule files for these SSNITs
                        st.markdown("##### 📁 Found in Schedule Files:")
                        schedule_findings = []
                        
                        for root, dirs, files in os.walk(main_folder_path):
                            for file in files:
                                if (file.lower().endswith('.xlsx') and 
                                    not file.startswith(('vlookup_', 'duplicate_ssnit_', '._', '~$'))):
                                    file_path = os.path.join(root, file)
                                    try:
                                        df = pd.read_excel(file_path)
                                        if 'ssnit' not in df.columns:
                                            st.warning(f"⚠️ No SSNIT column in {file}")
                                            continue
                                            
                                        df['ssnit'] = df['ssnit'].astype(str).str.strip()
                                        ssnit_in_file = set(df[df['ssnit'].isin(ssnit_set)]['ssnit'])
                                        
                                        if len(ssnit_in_file) > 1:
                                            schedule_findings.append({
                                                'File': file,
                                                'SSNITs Found': ', '.join(ssnit_in_file)
                                            })
                                            
                                    except Exception as e:
                                        st.error(f"Error processing {file}: {str(e)}")
                        
                        if schedule_findings:
                            st.table(pd.DataFrame(schedule_findings))
                        else:
                            st.info("No multiple SSNITs found in schedule files")
                
                # Export option
                if st.button("Export Results"):
                    export_path = os.path.join(main_folder_path, f"duplicate_ssnit_{company_name}.xlsx")
                    duplicates.to_excel(export_path, index=False)
                    st.success(f"Results exported to: {export_path}")
                
            else:
                st.success("No duplicate names found!")
                
        except Exception as e:
            st.error(f"Error processing duplicates: {str(e)}")

    def multiple_account_names(self):
        """Handle Multiple Account Names functionality"""
        st.header("🔍 Check Duplicate Account Numbers")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Get main company folder path
            main_folder_path = self.utils.select_folder("multiple_accounts")
            if main_folder_path:
                company_name = Path(main_folder_path).name
                st.success(f"Selected company folder: {company_name}")
        
        with col2:
            # Check for VLOOKUP file
            if main_folder_path:
                vlookup_path = Path(main_folder_path) / f"vlookup_{company_name}.xlsx"
                if vlookup_path.exists():
                    st.success(f"Found VLOOKUP file: {vlookup_path.name}")
                else:
                    st.error("VLOOKUP file not found! Please run the VLOOKUP process first.")
                    return
        
        if main_folder_path and st.button("Check for Duplicates", type="primary"):
            try:
                # Read VLOOKUP file
                df = pd.read_excel(vlookup_path)
                
                # Clean and standardize account numbers and names
                df['Accountno'] = df['Accountno'].astype(str).str.strip().str.upper()
                df['Full_Name'] = (df['Surname'].fillna('') + ' ' + 
                                 df['First_Name'].fillna('') + ' ' + 
                                 df['Other_Names'].fillna('')).str.strip().str.upper()
                
                # Find duplicates by account number
                account_duplicates = df[df.duplicated(subset='Accountno', keep=False)].sort_values('Accountno')
                
                # Find duplicates by full name
                name_duplicates = df[df.duplicated(subset='Full_Name', keep=False)].sort_values('Full_Name')
                
                st.markdown("---")
                
                # Display Account Number Duplicates
                st.subheader("🔢 Duplicate Account Numbers")
                if not account_duplicates.empty:
                    st.warning(f"Found {len(account_duplicates) // 2} duplicate Account Numbers!")
                    
                    # Group by Account number for better display
                    for account in account_duplicates['Accountno'].unique():
                        with st.expander(f"Account Number: {account}"):
                            account_records = account_duplicates[account_duplicates['Accountno'] == account]
                            
                            # Display with formatted columns
                            st.dataframe(
                                account_records[['Accountno', 'Surname', 'First_Name', 'Other_Names', 'Ssnit']],
                                column_config={
                                    "Accountno": "Account Number",
                                    "Surname": "Surname",
                                    "First_Name": "First Name",
                                    "Other_Names": "Other Names",
                                    "Ssnit": "SSNIT Number"
                                },
                                hide_index=True
                            )
                else:
                    st.success("No duplicate Account Numbers found!")
                
                st.markdown("---")
                
                # Display Name Duplicates
                st.subheader("👥 Duplicate Names")
                if not name_duplicates.empty:
                    st.warning(f"Found {len(name_duplicates) // 2} duplicate Names!")
                    
                    # Group by Full Name for better display
                    for name in name_duplicates['Full_Name'].unique():
                        with st.expander(f"Name: {name}"):
                            name_records = name_duplicates[name_duplicates['Full_Name'] == name]
                            
                            # Display with formatted columns
                            st.dataframe(
                                name_records[['Full_Name', 'Accountno', 'Ssnit']],
                                column_config={
                                    "Full_Name": "Full Name",
                                    "Accountno": "Account Number",
                                    "Ssnit": "SSNIT Number"
                                },
                                hide_index=True
                            )
                else:
                    st.success("No duplicate Names found!")
                
                # Export option if any duplicates found
                if not account_duplicates.empty or not name_duplicates.empty:
                    st.markdown("---")
                    if st.button("Export Duplicates"):
                        export_path = os.path.join(main_folder_path, f"duplicate_analysis_{company_name}.xlsx")
                        with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                            if not account_duplicates.empty:
                                account_duplicates.to_excel(writer, sheet_name='Account_Duplicates', index=False)
                            if not name_duplicates.empty:
                                name_duplicates.to_excel(writer, sheet_name='Name_Duplicates', index=False)
                        st.success(f"Exported duplicates to: {export_path}")
                    
                    # Display summary statistics
                    st.markdown("---")
                    st.subheader("📊 Summary")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Records", len(df))
                    with col2:
                        st.metric("Duplicate Accounts", len(account_duplicates) // 2 if not account_duplicates.empty else 0)
                    with col3:
                        st.metric("Duplicate Names", len(name_duplicates) // 2 if not name_duplicates.empty else 0)
                
            except Exception as e:
                st.error(f"Error checking for duplicates: {str(e)}")
                st.error("Please ensure the VLOOKUP file has the correct column names: Accountno, Surname, First_Name, Other_Names, Ssnit")

    def find_my_schedule(self):
        """Handle Find My Schedule functionality"""
        st.header("🔍 Find My Schedule")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Get main company folder path
            main_folder_path = self.utils.select_folder("find_schedule")
            if main_folder_path:
                company_name = Path(main_folder_path).name
                st.success(f"Selected company folder: {company_name}")
                
                # Get all schedule folders
                schedule_folders = [d for d in os.listdir(main_folder_path) 
                                  if os.path.isdir(os.path.join(main_folder_path, d))]
                
                if not schedule_folders:
                    st.error("No subfolders found in the company folder!")
                    return
        
        with col2:
            if main_folder_path:
                # SSNIT number input
                ssnit_number = st.text_input(
                    "Enter SSNIT Number to find:",
                    help="Enter the SSNIT number you want to locate"
                ).strip()
                
                # Search type selection
                search_type = st.radio(
                    "Select where to search:",
                    ["VLOOKUP File", "Schedule Files", "Both"],
                    horizontal=True
                )
                
                if search_type in ["Schedule Files", "Both"]:
                    folder_options = ["Search All Folders"] + schedule_folders
                    selected_option = st.selectbox(
                        "Select schedule folder:",
                        folder_options,
                        help="Choose a specific folder or search all folders"
                    )
        
        if main_folder_path and ssnit_number:
            if st.button("Find SSNIT Number", type="primary"):
                try:
                    found_records = []
                    
                    # Search in VLOOKUP file if selected
                    if search_type in ["VLOOKUP File", "Both"]:
                        vlookup_path = Path(main_folder_path) / f"vlookup_{company_name}.xlsx"
                        if vlookup_path.exists():
                            vlookup_df = pd.read_excel(vlookup_path)
                            vlookup_df['Ssnit'] = vlookup_df['Ssnit'].astype(str).str.strip()
                            
                            vlookup_matches = vlookup_df[vlookup_df['Ssnit'] == ssnit_number]
                            
                            if not vlookup_matches.empty:
                                st.success("📋 Found in VLOOKUP File:")
                                st.dataframe(
                                    vlookup_matches,
                                    column_config={
                                        "Ssnit": "SSNIT Number",
                                        "Surname": "Surname",
                                        "First_Name": "First Name",
                                        "Other_Names": "Other Names",
                                        "Accountno": "Account Number"
                                    },
                                    hide_index=True
                                )
                        else:
                            st.warning("VLOOKUP file not found!")
                    
                    # Search in schedule files if selected
                    if search_type in ["Schedule Files", "Both"]:
                        st.markdown("---")
                        st.subheader("📁 Schedule Files Search")
                        
                        folders_to_search = schedule_folders if selected_option == "Search All Folders" else [selected_option]
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        for folder_idx, folder in enumerate(folders_to_search):
                            folder_path = os.path.join(main_folder_path, folder)
                            status_text.text(f"Searching in folder: {folder}")
                            
                            files = [f for f in os.listdir(folder_path) 
                                    if f.lower().endswith('.xlsx') and 
                                    not f.startswith((f"vlookup_{company_name}", 
                                                    f"duplicate_ssnit_{company_name}",
                                                    "._", "~$", "duplicate_"))]
                            
                            for file in files:
                                file_path = os.path.join(folder_path, file)
                                try:
                                    df = pd.read_excel(file_path)
                                    
                                    if 'ssnit' not in df.columns:
                                        continue
                                        
                                    df['ssnit'] = df['ssnit'].astype(str).str.strip()
                                    matches = df[df['ssnit'] == ssnit_number]
                                    
                                    if not matches.empty:
                                        for _, row in matches.iterrows():
                                            record = {
                                                'Folder': folder,
                                                'File': file,
                                                'SSNIT': ssnit_number
                                            }
                                            
                                            if 'name' in df.columns:
                                                record['Name'] = row['name']
                                            if 'salary' in df.columns:
                                                # Clean salary value by removing spaces and commas before converting to float
                                                salary_str = str(row['salary']).strip().replace(',', '').replace(' ', '')
                                                salary_value = float(salary_str)
                                                record['Salary'] = salary_value
                                            
                                            found_records.append(record)
                                
                                except Exception as e:
                                    st.error(f"Error processing {folder}/{file}: {str(e)}")
                            
                            progress_bar.progress((folder_idx + 1) / len(folders_to_search))
                        
                        progress_bar.empty()
                        status_text.empty()
                        
                        # Display schedule search results
                        if found_records:
                            st.success(f"Found {len(found_records)} occurrence(s) in schedule files:")
                            results_df = pd.DataFrame(found_records)
                            st.dataframe(
                                results_df,
                                column_config={
                                    "Folder": "Schedule Folder",
                                    "File": "File Name",
                                    "SSNIT": "SSNIT Number",
                                    "Name": "Employee Name",
                                    "Salary": st.column_config.NumberColumn(
                                        "Salary",
                                        help="Employee Salary",
                                        format="%.2f"
                                    )
                                },
                                hide_index=True
                            )
                        else:
                            st.info("No matches found in schedule files")
                    
                    # Export option if any results found
                    if (search_type in ["VLOOKUP File", "Both"] and not vlookup_matches.empty) or \
                       (search_type in ["Schedule Files", "Both"] and found_records):
                        st.markdown("---")
                        if st.button("Export Results"):
                            export_path = os.path.join(main_folder_path, f"ssnit_search_{ssnit_number}.xlsx")
                            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                                if search_type in ["VLOOKUP File", "Both"] and not vlookup_matches.empty:
                                    vlookup_matches.to_excel(writer, sheet_name='VLOOKUP_Results', index=False)
                                if search_type in ["Schedule Files", "Both"] and found_records:
                                    pd.DataFrame(found_records).to_excel(writer, sheet_name='Schedule_Results', index=False)
                            st.success(f"Exported results to: {export_path}")
                
                except Exception as e:
                    st.error(f"Error during search: {str(e)}")

    def validation(self):
        """Handle Validation functionality"""
        st.header("Validation")
        
        # Get main company folder path
        main_folder_path = self.utils.select_folder("validation")
        if not main_folder_path:
            st.warning("Please select a folder first!")
            return
        
        company_name = Path(main_folder_path).name
        st.success(f"Selected company folder: {company_name}")
        
        # Check for required files
        vlookup_path = Path(main_folder_path) / f"vlookup_{company_name}.xlsx"
        
        # Add master report file upload
        st.subheader("Upload Master Report")
        master_file = st.file_uploader(
            "Upload master report (almighty report)",
            type=['xlsx'],
            key="master_upload"
        )
        
        if not all([master_file, vlookup_path.exists()]):
            if not master_file:
                st.warning("Please upload the master report file")
            if not vlookup_path.exists():
                st.error("VLOOKUP file not found! Please run the VLOOKUP process first.")
            return
        
        try:
            # Read and process files
            master_df = pd.read_excel(master_file)
            vlookup_db = pd.read_excel(vlookup_path)
            
            # Clean and standardize the data
            master_df['Ssnit'] = master_df['Ssnit'].astype(str).str.strip()
            vlookup_db['Ssnit'] = vlookup_db['Ssnit'].astype(str).str.strip()
            
            # Create dictionaries for lookups
            vlookup_dict = {}
            for _, row in vlookup_db.iterrows():
                ssnit = str(row['Ssnit']).strip()
                vlookup_dict[ssnit] = {
                    'accountno': str(row['Accountno']).strip() if pd.notna(row['Accountno']) else None,
                    'surname': row['Surname'] if pd.notna(row['Surname']) else None,
                    'first_name': row['First_Name'] if pd.notna(row['First_Name']) else None,
                    'other_name': row['Other_Names'] if pd.notna(row['Other_Names']) else None
                }
            
            # Add master data for missing entries
            for _, row in master_df.iterrows():
                ssnit = str(row['Ssnit']).strip()
                if ssnit not in vlookup_dict:
                    vlookup_dict[ssnit] = {
                        'accountno': str(row['Client Account Number']).strip() if pd.notna(row['Client Account Number']) else None,
                        'surname': row['Surname'] if pd.notna(row['Surname']) else None,
                        'first_name': row['First Name'] if pd.notna(row['First Name']) else None,
                        'other_name': row['Other Names'] if pd.notna(row['Other Names']) else None
                    }
                else:
                    # Fill in missing data from master
                    entry = vlookup_dict[ssnit]
                    if not entry['accountno']:
                        entry['accountno'] = str(row['Client Account Number']).strip() if pd.notna(row['Client Account Number']) else None
                    if not entry['surname']:
                        entry['surname'] = row['Surname']
                    if not entry['first_name']:
                        entry['first_name'] = row['First Name']
                    if not entry['other_name']:
                        entry['other_name'] = row['Other Names']
            
            # Process files
            files_to_process = []
            for root, _, files in os.walk(main_folder_path):
                for file in files:
                    if (file.lower().endswith('.xlsx') and 
                        not file.startswith(('vlookup_', 'duplicate_', '._', '~$'))):
                        files_to_process.append({
                            'path': os.path.join(root, file),
                            'name': file,
                            'status': 'Pending'
                        })
            
            if not files_to_process:
                st.warning("No files found to process!")
                return
            
            # Display files
            st.subheader("Files to Process")
            files_df = pd.DataFrame(files_to_process)
            st.dataframe(files_df[['name', 'status']], hide_index=True)
            
            if st.button("Start Validation", key="start_validation", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, file_info in enumerate(files_to_process):
                    try:
                        file_path = file_info['path']
                        status_text.text(f"Processing: {file_info['name']}")
                        
                        # Process file
                        df = pd.read_excel(file_path)
                        df['ssnit'] = df['ssnit'].astype(str).str.strip()
                        
                        # Map SSNIT data
                        for ssnit_num in df['ssnit'].unique():
                            mask = df['ssnit'] == ssnit_num
                            if ssnit_num in vlookup_dict:
                                data = vlookup_dict[ssnit_num]
                                for field, value in data.items():
                                    if value:  # Only update if value exists
                                        df.loc[mask, field] = value
                        
                        # Calculate tiers
                        df['salary'] = pd.to_numeric(df['salary'].astype(str).str.replace(',', '').str.strip(), errors='coerce')
                        df['tier1'] = 0
                        df['tier2'] = df['salary'] * 0.05
                        
                        # Save processed file
                        output_columns = ['accountno', 'surname', 'first_name', 'other_name', 'ssnit', 'tier1', 'tier2']
                        df[output_columns].to_excel(file_path, index=False)
                        
                        files_df.loc[idx, 'status'] = 'Processed'
                        
                    except Exception as e:
                        files_df.loc[idx, 'status'] = 'Failed'
                        st.error(f"Error processing {file_info['name']}: {str(e)}")
                    
                    progress_bar.progress((idx + 1) / len(files_to_process))
                
                progress_bar.empty()
                status_text.empty()
                
                st.success("Validation completed!")
                st.dataframe(files_df[['name', 'status']], hide_index=True)
                
        except Exception as e:
            st.error(f"Error during validation: {str(e)}")

    def append_total(self):
        """Handle Append Total functionality""" 
        st.header("📊 Append Total")
        
        # Get main company folder path
        main_folder_path = self.utils.select_folder("append_total")
        if not main_folder_path:
            st.warning("Please enter the company folder path")
            return
        
        company_name = Path(main_folder_path).name
        st.success(f"Selected company folder: {company_name}")
        
        # Create column for folder selection
        col1, _ = st.columns(2)
        
        with col1:
            # Get all schedule folders
            schedule_folders = [d for d in os.listdir(main_folder_path) 
                              if os.path.isdir(os.path.join(main_folder_path, d))]
            
            if not schedule_folders:
                st.error("No subfolders found in the company folder!")
                return
            
            # Add option to process all folders
            folder_options = ["Process All Folders"] + schedule_folders
            selected_option = st.selectbox(
                "Select schedule folder:",
                folder_options,
                help="Choose a specific folder or process all folders"
            )
            
            # Determine which folders to process
            folders_to_process = schedule_folders if selected_option == "Process All Folders" else [selected_option]
            
            st.info(f"Will process: {', '.join(folders_to_process)}")
        
        # Files Preview Section
        st.markdown("---")
        st.subheader("📁 Files to be Processed")
        
        try:
            # Get list of all files to process
            all_files = []
            for folder in folders_to_process:
                folder_path = os.path.join(main_folder_path, folder)
                for file in os.listdir(folder_path):
                    if (file.endswith('.xlsx') and 
                        not file.startswith(('vlookup_', 'duplicate_ssnit_', '._', '~$')) and
                        '_' not in file):  # Skip already processed files
                        all_files.append({
                            'Folder': folder,
                            'File Name': file,
                            'Status': 'Pending'
                        })
            
            if all_files:
                # Create DataFrame for status tracking
                files_df = pd.DataFrame(all_files)
                
                # Display files in a nice table
                st.dataframe(
                    files_df,
                    column_config={
                        "Folder": "Schedule Folder",
                        "File Name": "Schedule Files",
                        "Status": st.column_config.Column(
                            "Status",
                            help="Processing status of each file",
                            width="medium"
                        )
                    },
                    hide_index=True
                )
                
                # Add Start Append Total button
                if st.button("Start Append Total", type="primary"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    status_container = st.empty()
                    processed_count = 0
                    
                    for idx, file_info in enumerate(all_files):
                        folder = file_info['Folder']
                        file = file_info['File Name']
                        folder_path = os.path.join(main_folder_path, folder)
                        file_path = os.path.join(folder_path, file)
                        
                        status_text.text(f"Processing: {folder}/{file}")
                        
                        try:
                            # Read the Excel file
                            df = pd.read_excel(file_path)
                            
                            # Calculate total tier2
                            if 'tier2' not in df.columns:
                                raise ValueError(f"tier2 column not found in {file}")
                            
                            total_tier2 = df['tier2'].sum()
                            total_tier2_formatted = f"{total_tier2:.2f}"
                            
                            # Create new filename with total amount
                            filename_without_extension = os.path.splitext(file)[0]
                            extension = '.xlsx'
                            new_filename = f"{filename_without_extension}_{total_tier2_formatted}{extension}"
                            new_file_path = os.path.join(folder_path, new_filename)
                            
                            # Rename the file
                            os.rename(file_path, new_file_path)
                            
                            files_df.loc[idx, 'Status'] = 'Processed ✅'
                            processed_count += 1
                            
                            # Display data in expander
                            with st.expander(f"View Results: {folder}/{new_filename}", expanded=False):
                                st.markdown("##### Sample of Processed Data")
                                st.dataframe(
                                    df.head(),
                                    column_config={
                                        "accountno": st.column_config.Column(
                                            "Account Number",
                                            width="medium"
                                        ),
                                        "surname": st.column_config.Column(
                                            "Surname",
                                            width="medium"
                                        ),
                                        "first_name": st.column_config.Column(
                                            "First Name",
                                            width="medium"
                                        ),
                                        "other_name": st.column_config.Column(
                                            "Other Name",
                                            width="medium"
                                        ),
                                        "ssnit": st.column_config.Column(
                                            "SSNIT Number",
                                            width="medium"
                                        ),
                                        "tier1": st.column_config.NumberColumn(
                                            "Tier 1",
                                            format="₵%.2f",
                                            width="medium"
                                        ),
                                        "tier2": st.column_config.NumberColumn(
                                            "Tier 2",
                                            format="₵%.2f",
                                            width="medium"
                                        )
                                    },
                                    hide_index=True
                                )
                                
                                # Display file statistics
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.metric("Total Records", len(df))
                                with col2:
                                    st.metric("Total Tier 2", f"₵{total_tier2:,.2f}")
                        
                        except Exception as e:
                            files_df.loc[idx, 'Status'] = 'Failed ❌'
                            st.error(f"Error processing {folder}/{file}: {str(e)}")
                        
                        # Update status display
                        status_container.dataframe(
                            files_df,
                            column_config={
                                "Folder": "Schedule Folder",
                                "File Name": "Schedule Files",
                                "Status": st.column_config.Column(
                                    "Status",
                                    help="Processing status of each file",
                                    width="medium"
                                )
                            },
                            hide_index=True
                        )
                        
                        progress_bar.progress((idx + 1) / len(files_df))
                    
                    progress_bar.empty()
                    status_text.empty()
                    
                    # Show completion message
                    if processed_count > 0:
                        st.success(f"""
                            Total Append completed successfully!
                            - Files processed: {processed_count} out of {len(all_files)}
                            - Success rate: {(processed_count/len(all_files))*100:.1f}%
                        """)
                        
                        # Display processed files
                        st.subheader("Processed Files")
                        processed_files = files_df[files_df['Status'] == 'Processed ✅']
                        if not processed_files.empty:
                            for _, row in processed_files.iterrows():
                                st.markdown(f"- ✅ {row['Folder']}/{row['File Name']}")
            else:
                st.warning("No files found to process!")
                
        except Exception as e:
            st.error(f"Error accessing directories: {str(e)}")

    def find_duplicates(self):
        """Handle Find Duplicates functionality"""
        st.header("🔍 Find Duplicates")
        
        # Get main company folder path
        main_folder_path = self.utils.select_folder("find_duplicates")
        if not main_folder_path:
            st.warning("Please enter the company folder path")
            return
        
        company_name = Path(main_folder_path).name
        st.success(f"Selected company folder: {company_name}")
        
        # Look for vlookup file
        vlookup_path = Path(main_folder_path) / f"vlookup_{company_name}.xlsx"
        if not vlookup_path.exists():
            st.error("VLOOKUP file not found! Please run the VLOOKUP process first.")
            return
        
        try:
            # Read the VLOOKUP file
            df = pd.read_excel(vlookup_path)
            
            # Create a copy of the DataFrame for case-insensitive comparison
            df_compare = df.copy()
            
            # Combine name columns and standardize
            df_compare['full_name'] = df_compare.apply(
                lambda row: ' '.join(filter(None, [
                    str(row['Surname']).strip(),
                    str(row['First_Name']).strip(),
                    str(row['Other_Names']).strip()
                ])), axis=1
            ).str.upper()
            
            # Find duplicates based on SSNIT numbers
            ssnit_duplicates = df[df.duplicated(subset=['Ssnit'], keep=False)].sort_values('Ssnit')
            
            # Find duplicates based on combined names
            duplicate_mask = df_compare.duplicated(subset=['full_name'], keep=False)
            name_duplicates = df[duplicate_mask].sort_values(['Surname', 'First_Name', 'Other_Names'])
            
            # Display results
            if not ssnit_duplicates.empty:
                st.subheader("🔍 Duplicate SSNIT Numbers Found")
                st.dataframe(
                    ssnit_duplicates,
                    column_config={
                        "Accountno": "Account Number",
                        "Surname": "Surname",
                        "First_Name": "First Name",
                        "Other_Names": "Other Names",
                        "Ssnit": "SSNIT Number",
                        "Accountno2": "Account Number 2"
                    },
                    hide_index=True
                )
            
            if not name_duplicates.empty:
                st.subheader("🔍 Duplicate Names Found")
                st.dataframe(
                    name_duplicates,
                    column_config={
                        "Accountno": "Account Number",
                        "Surname": "Surname",
                        "First_Name": "First Name",
                        "Other_Names": "Other Names",
                        "Ssnit": "SSNIT Number",
                        "Accountno2": "Account Number 2"
                    },
                    hide_index=True
                )
                
            if ssnit_duplicates.empty and name_duplicates.empty:
                st.success("No duplicates found! ✨")
                
        except Exception as e:
            st.error(f"Error processing duplicates: {str(e)}")

def standardize_name(name):
    """Standardize name by sorting words alphabetically"""
    return ' '.join(sorted(str(name).upper().split()))

def create_comprehensive_mapping(vlookup_df, master_df):
    """Create a comprehensive mapping using both VLOOKUP and master data"""
    mapping = {}
    
    # Clean and standardize SSNIT numbers in both dataframes
    vlookup_df['Ssnit'] = vlookup_df['Ssnit'].astype(str).str.strip()
    master_df['Ssnit'] = master_df['Ssnit'].astype(str).str.strip()
    
    # First populate from VLOOKUP (primary source)
    for _, row in vlookup_df.iterrows():
        ssnit = row['Ssnit']
        if pd.notna(ssnit) and str(ssnit).strip():
            mapping[ssnit] = {
                'accountno': str(row['Accountno']).strip() if pd.notna(row['Accountno']) else None,
                'surname': str(row['Surname']).strip() if pd.notna(row['Surname']) else None,
                'first_name': str(row['First_Name']).strip() if pd.notna(row['First_Name']) else None,
                'other_name': str(row['Other_Names']).strip() if pd.notna(row['Other_Names']) else None,
                'source': 'VLOOKUP'
            }
    
    # Supplement with master data where missing
    for _, row in master_df.iterrows():
        ssnit = row['Ssnit']
        if pd.notna(ssnit) and str(ssnit).strip():
            if ssnit not in mapping:
                # Add new entry from master
                mapping[ssnit] = {
                    'accountno': str(row['Client Account Number']).strip() if pd.notna(row['Client Account Number']) else None,
                    'surname': str(row['Surname']).strip() if pd.notna(row['Surname']) else None,
                    'first_name': str(row['First Name']).strip() if pd.notna(row['First Name']) else None,
                    'other_name': str(row['Other Names']).strip() if pd.notna(row['Other Names']) else None,
                    'source': 'Master'
                }
            else:
                # Fill in missing data from master
                entry = mapping[ssnit]
                if not entry['accountno']:
                    entry['accountno'] = str(row['Client Account Number']).strip() if pd.notna(row['Client Account Number']) else None
                if not entry['surname']:
                    entry['surname'] = str(row['Surname']).strip() if pd.notna(row['Surname']) else None
                if not entry['first_name']:
                    entry['first_name'] = str(row['First Name']).strip() if pd.notna(row['First Name']) else None
                if not entry['other_name']:
                    entry['other_name'] = str(row['Other Names']).strip() if pd.notna(row['Other Names']) else None
    
    return mapping

def process_dataframe(df, vlookup_dict, master_dict):
    """Process DataFrame using consistent SSNIT-Account mapping"""
    # Get the consistent mapping
    ssnit_mapping = create_comprehensive_mapping(vlookup_dict, master_dict)
    
    # Convert to string and clean SSNIT numbers
    df['ssnit'] = df['ssnit'].astype(str).str.strip()
    
    # Create result DataFrame
    result_df = pd.DataFrame(index=df.index)
    result_df['ssnit'] = df['ssnit']
    
    # Map the consistent account numbers and other fields
    for field in ['accountno', 'surname', 'first_name', 'other_name']:
        result_df[field] = df['ssnit'].map(
            lambda x: ssnit_mapping.get(x, {}).get(field, '#N/A' if field == 'accountno' else '')
        )
    
    # Handle salary and tiers
    df['salary'] = pd.to_numeric(df['salary'].astype(str).str.replace(',', '').str.strip(), errors='coerce')
    result_df['tier1'] = 0
    result_df['tier2'] = df['salary'] * 0.05
    
    return result_df[['accountno', 'surname', 'first_name', 'other_name', 'ssnit', 'tier1', 'tier2']]

def check_and_standardize_accounts(folder_path, vlookup_dict, master_dict):
    """Check and standardize account numbers for each SSNIT before validation"""
    # Create consistent SSNIT mapping
    ssnit_mapping = create_comprehensive_mapping(vlookup_dict, master_dict)
    
    if ssnit_mapping is None:
        return False
    
    modified_files = []
    
    # Process each schedule file
    for root, _, files in os.walk(folder_path):
        for file in files:
            if (file.lower().endswith('.xlsx') and 
                not file.startswith(('vlookup_', 'duplicate_', '._', '~$'))):
                file_path = os.path.join(root, file)
                try:
                    # Read file
                    df = pd.read_excel(file_path)
                    if 'ssnit' not in df.columns:
                        continue
                    
                    # Clean SSNIT numbers
                    df['ssnit'] = df['ssnit'].astype(str).str.strip()
                    
                    # Find records that need standardization
                    changes_made = False
                    for idx, row in df.iterrows():
                        ssnit = row['ssnit']
                        if ssnit in ssnit_mapping:
                            standard_account = ssnit_mapping[ssnit]['accountno']
                            if row.get('accountno') != standard_account:
                                df.at[idx, 'accountno'] = standard_account
                                changes_made = True
                    
                    # Save changes if any made
                    if changes_made:
                        df.to_excel(file_path, index=False)
                        modified_files.append(file)
                        
                except Exception as e:
                    st.error(f"Error processing {file}: {str(e)}")
    
    if modified_files:
        st.success(f"Updated {len(modified_files)} files with standardized account numbers")
        return True
    
    return True

def check_individual_schedule_duplicates(folder_path):
    """Check for duplicates within each schedule file and standardize them"""
    st.subheader("🔍 Pre-validation Duplicate Check")
    
    duplicates_found = False
    duplicates_by_file = {}
    
    for root, _, files in os.walk(folder_path):
        for file in files:
            if (file.lower().endswith('.xlsx') and 
                not file.startswith(('vlookup_', 'duplicate_', '._', '~$'))):
                file_path = os.path.join(root, file)
                try:
                    df = pd.read_excel(file_path)
                    if 'ssnit' not in df.columns:
                        continue
                        
                    df['ssnit'] = df['ssnit'].astype(str).str.strip()
                    duplicates = df[df.duplicated(subset='ssnit', keep=False)]
                    
                    if not duplicates.empty:
                        duplicates_found = True
                        duplicates_by_file[file] = duplicates
                        
                        # Show duplicates in expandable section
                        with st.expander(f"📄 Found duplicates in {file}:"):
                            st.dataframe(duplicates)
                            
                            # Group by SSNIT number
                            for ssnit in duplicates['ssnit'].unique():
                                ssnit_entries = duplicates[duplicates['ssnit'] == ssnit]
                                st.markdown(f"##### SSNIT: {ssnit}")
                                
                                # Take the first account number for standardization
                                primary_account = ssnit_entries.iloc[0]['accountno']
                                
                                # Update all entries with this SSNIT to use the same account
                                mask = df['ssnit'] == ssnit
                                df.loc[mask, 'accountno'] = primary_account
                                
                                st.code(f"""
                                Standardized account number to: {primary_account}
                                """)
                            
                            # Save changes back to file
                            df.to_excel(file_path, index=False)
                            
                except Exception as e:
                    st.error(f"Error checking {file}: {str(e)}")
    
    if duplicates_found:
        st.warning("🔄 Duplicates were found and standardized")
        if st.button("Continue with Validation", type="primary"):
            return True
        return False
    else:
        st.success("✅ No duplicates found in schedule files")
        if st.button("Continue with Validation", type="primary"):
            return True
        return False

def create_comprehensive_mapping(vlookup_df, master_df):
    """Create a comprehensive mapping using both VLOOKUP and master data"""
    mapping = {}
    
    # Clean and standardize SSNIT numbers in both dataframes
    vlookup_df['Ssnit'] = vlookup_df['Ssnit'].astype(str).str.strip().upper()
    master_df['Ssnit'] = master_df['Ssnit'].astype(str).str.strip().upper()
    
    # Track unmapped SSNITs
    unmapped_ssnits = set()
    
    # First populate from VLOOKUP (primary source)
    for _, row in vlookup_df.iterrows():
        ssnit = row['Ssnit']
        if pd.notna(ssnit) and str(ssnit).strip() and ssnit not in ['NAN', 'NONE', '']:
            mapping[ssnit] = {
                'accountno': str(row['Accountno']).strip() if pd.notna(row['Accountno']) else '',
                'surname': str(row['Surname']).strip() if pd.notna(row['Surname']) else '',
                'first_name': str(row['First_Name']).strip() if pd.notna(row['First_Name']) else '',
                'other_name': str(row['Other_Names']).strip() if pd.notna(row['Other_Names']) else '',
                'source': 'VLOOKUP'
            }
    
    # Supplement with master data
    for _, row in master_df.iterrows():
        ssnit = row['Ssnit']
        if pd.notna(ssnit) and str(ssnit).strip() and ssnit not in ['NAN', 'NONE', '']:
            if ssnit not in mapping:
                mapping[ssnit] = {
                    'accountno': str(row['Client Account Number']).strip() if pd.notna(row['Client Account Number']) else '',
                    'surname': str(row['Surname']).strip() if pd.notna(row['Surname']) else '',
                    'first_name': str(row['First Name']).strip() if pd.notna(row['First Name']) else '',
                    'other_name': str(row['Other Names']).strip() if pd.notna(row['Other Names']) else '',
                    'source': 'Master'
                }
            else:
                # Fill in missing data from master
                entry = mapping[ssnit]
                if not entry['accountno']:
                    entry['accountno'] = str(row['Client Account Number']).strip() if pd.notna(row['Client Account Number']) else ''
    
    return mapping

def process_schedule_files(folder_path, ssnit_mapping):
    """Process schedule files with improved validation"""
    modified_files = []
    unmapped_records = []
    
    for root, _, files in os.walk(folder_path):
        for file in files:
            if (file.lower().endswith('.xlsx') and 
                not file.startswith(('vlookup_', 'duplicate_', '._', '~$'))):
                file_path = os.path.join(root, file)
                try:
                    df = pd.read_excel(file_path)
                    if 'ssnit' not in df.columns:
                        continue
                    
                    # Clean SSNIT numbers
                    df['ssnit'] = df['ssnit'].astype(str).str.strip().upper()
                    
                    # Track changes and unmapped SSNITs
                    changes_made = False
                    file_unmapped = []
                    
                    # Process each row
                    for idx, row in df.iterrows():
                        ssnit = row['ssnit']
                        if ssnit in ['NAN', 'NONE', '']:
                            continue
                            
                        if ssnit in ssnit_mapping:
                            # Update only if data exists in mapping
                            data = ssnit_mapping[ssnit]
                            if data['accountno']:  # Only update if we have a valid account number
                                df.at[idx, 'accountno'] = data['accountno']
                                df.at[idx, 'surname'] = data['surname']
                                df.at[idx, 'first_name'] = data['first_name']
                                df.at[idx, 'other_name'] = data['other_name']
                                changes_made = True
                        else:
                            # Track unmapped SSNITs
                            file_unmapped.append({
                                'file': file,
                                'ssnit': ssnit,
                                'row': idx + 1
                            })
                    
                    # Save changes if any were made
                    if changes_made:
                        df.to_excel(file_path, index=False)
                        modified_files.append(file)
                    
                    # Add unmapped records
                    unmapped_records.extend(file_unmapped)
                    
                except Exception as e:
                    st.error(f"Error processing {file}: {str(e)}")
    
    # Display results
    if modified_files:
        st.success(f"✅ Updated {len(modified_files)} files with standardized data")
    
    if unmapped_records:
        st.warning("⚠️ Found unmapped SSNIT numbers:")
        for record in unmapped_records:
            st.write(f"File: {record['file']}, Row: {record['row']}, SSNIT: {record['ssnit']}")
    
    return modified_files, unmapped_records
        
    # Add this at the bottom of the file:
if __name__ == "__main__":
    dashboard = Dashboard()
    dashboard.run()