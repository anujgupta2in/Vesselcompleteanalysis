import pandas as pd
import numpy as np
import re

class BWTSProcessor:
    def __init__(self):
        """Initialize BWTSProcessor with component list."""
        self.components = [
            'BWTS Control Panel',
            'BWTS Filter Unit',
            'BWTS UV System',
            'BWTS Reactor',
            'BWTS Pump',
            'BWTS Neutralizing Unit',
            'BWTS Chemical Dosing Unit',
            'BWTS Ballast Tank Sensor',
            'BWTS Monitoring System',
            'BWTS Sample Point'
        ]
    
    def extract_running_hours(self, data):
        """Extract running hours for BWTS."""
        try:
            # Filter data for BWTS with more flexible patterns
            bwts_patterns = ['Ballast Water Treatment Plant', 'BWTS', 'Ballast Treatment']
            mask = data['Machinery Location'].str.contains('|'.join(bwts_patterns), case=False, na=False)
            bwts_data = data[mask].copy()
            print(f"Found {len(bwts_data)} BWTS records using patterns: {bwts_patterns}")
            
            if bwts_data.empty:
                # Return empty DataFrame with expected columns to avoid errors
                return pd.DataFrame(columns=['BWTS Unit', 'Running Hours'])
            
            # Initialize empty list to store running hours data
            running_hours_data = []
            
            # Get unique BWTS units
            def extract_bwts_number(location):
                if not isinstance(location, str):
                    return "Unknown BWTS Unit"
                    
                # Try multiple patterns to extract unit numbers
                bwts_patterns = [
                    r'Ballast Water Treatment Plant#?(\d+)',
                    r'BWTS#?(\d+)',
                    r'Ballast Treatment#?(\d+)'
                ]
                
                for pattern in bwts_patterns:
                    match = re.search(pattern, location, re.IGNORECASE)
                    if match:
                        return f"BWTS Unit #{match.group(1)}"
                
                # If no match, just return the location as is (with a safety check)
                return location[:100] if len(location) > 100 else location
            
            # Make sure there's a BWTS Unit column
            bwts_data['BWTS Unit'] = bwts_data['Machinery Location'].apply(extract_bwts_number)
            unique_units = bwts_data['BWTS Unit'].unique()
            
            # For each unit, get the running hours
            for unit in unique_units:
                unit_data = bwts_data[bwts_data['BWTS Unit'] == unit]
                
                # Get the running hours for this unit
                running_hours = 0  # Default to 0
                if 'Machinery Running Hours' in unit_data.columns:
                    # Convert to numeric, coercing errors to NaN
                    unit_data['Machinery Running Hours'] = pd.to_numeric(unit_data['Machinery Running Hours'], errors='coerce')
                    running_hours_values = unit_data['Machinery Running Hours'].dropna()
                    if not running_hours_values.empty:
                        running_hours = running_hours_values.iloc[0]
                
                # Add to our list with number format to avoid serialization issues
                running_hours_data.append({
                    'BWTS Unit': str(unit),  # Ensure string type
                    'Running Hours': float(running_hours) if pd.notna(running_hours) else 0.0  # Ensure numeric type
                })
            
            # Convert to DataFrame and ensure proper types
            result_df = pd.DataFrame(running_hours_data)
            if result_df.empty:
                return pd.DataFrame(columns=['BWTS Unit', 'Running Hours'])
                
            return result_df
            
        except Exception as e:
            print(f"Error extracting BWTS running hours: {str(e)}")
            # Return empty DataFrame with expected columns
            return pd.DataFrame(columns=['BWTS Unit', 'Running Hours'])
    
    def format_unit_data(self, unit_data):
        """Format unit data for display."""
        try:
            # Create a copy to avoid modifying original
            data_copy = unit_data.copy()
            
            # Ensure all entries are strings
            for col in data_copy.columns:
                data_copy[col] = data_copy[col].astype(str)
            
            return data_copy
            
        except Exception as e:
            print(f"Error formatting BWTS unit data: {str(e)}")
            return unit_data
    
    def process_job_code(self, data, job_codes, job_description, unit_pattern=r'(?:Ballast Water Treatment Plant|BWTS)#?(\d+)'):
        """Process job codes for BWTS."""
        try:
            # Filter data for BWTS with more flexible patterns
            bwts_patterns = ['Ballast Water Treatment Plant', 'BWTS', 'Ballast Treatment']
            mask = data['Machinery Location'].str.contains('|'.join(bwts_patterns), case=False, na=False)
            bwts_data = data[mask].copy()
            print(f"Found {len(bwts_data)} BWTS records in process_job_code using patterns: {bwts_patterns}")
            
            if bwts_data.empty:
                return pd.DataFrame()
                
            # Process the data
            bwts_data['Unit'] = bwts_data['Machinery Location'].str.extract(unit_pattern, expand=False)
            bwts_data['Unit'] = bwts_data['Unit'].fillna('General')
            bwts_data['Unit'] = 'Unit #' + bwts_data['Unit']
            
            # Create pivot table
            if 'Job Code' in bwts_data.columns and job_codes in bwts_data.columns and job_description in bwts_data.columns:
                pivot_df = pd.pivot_table(
                    bwts_data,
                    values='Job Code',
                    index=job_description,
                    columns='Unit',
                    aggfunc='count',
                    fill_value=0
                )
                
                # Reset index to convert to regular DataFrame
                pivot_df = pivot_df.reset_index()
                
                return pivot_df
            else:
                print(f"Required columns not found for BWTS job code processing.")
                return pd.DataFrame()
                
        except Exception as e:
            print(f"Error processing BWTS job codes: {str(e)}")
            return pd.DataFrame()
    
    def get_maintenance_data(self, data):
        """Get BWTS maintenance data."""
        try:
            # Filter data for BWTS with more flexible patterns
            bwts_patterns = ['Ballast Water Treatment Plant', 'BWTS', 'Ballast Treatment']
            mask = data['Machinery Location'].str.contains('|'.join(bwts_patterns), case=False, na=False)
            bwts_data = data[mask].copy()
            print(f"Found {len(bwts_data)} BWTS records in get_maintenance_data using patterns: {bwts_patterns}")
            
            if bwts_data.empty:
                return pd.DataFrame()
            
            # Select relevant columns
            if 'Job Code' in bwts_data.columns and 'Title' in bwts_data.columns and 'Calculated Due Date' in bwts_data.columns:
                maintenance_data = bwts_data[['Job Code', 'Title', 'Calculated Due Date', 'Machinery Location', 'Job Status']].copy()
                
                # Sort by due date
                if 'Calculated Due Date' in maintenance_data.columns:
                    maintenance_data = maintenance_data.sort_values(by='Calculated Due Date')
                
                return maintenance_data
            else:
                print("Required columns not found for BWTS maintenance data.")
                return pd.DataFrame()
                
        except Exception as e:
            print(f"Error getting BWTS maintenance data: {str(e)}")
            return pd.DataFrame()
    
    def analyze_components(self, data):
        """Analyze component presence for BWTS."""
        try:
            # Filter data for BWTS with more flexible patterns
            bwts_patterns = ['Ballast Water Treatment Plant', 'BWTS', 'Ballast Treatment']
            mask = data['Machinery Location'].str.contains('|'.join(bwts_patterns), case=False, na=False)
            bwts_data = data[mask].copy()
            print(f"Found {len(bwts_data)} BWTS records in analyze_components using patterns: {bwts_patterns}")
            
            if bwts_data.empty:
                return pd.DataFrame()
                
            # Extract BWTS unit numbers
            def extract_bwts_number(location):
                # Try multiple patterns to extract unit numbers
                bwts_patterns = [
                    r'Ballast Water Treatment Plant#?(\d+)',
                    r'BWTS#?(\d+)',
                    r'Ballast Treatment#?(\d+)'
                ]
                
                for pattern in bwts_patterns:
                    match = re.search(pattern, location, re.IGNORECASE)
                    if match:
                        return int(match.group(1))
                
                return 0  # default value for general items
                
            bwts_data['Unit Number'] = bwts_data['Machinery Location'].apply(extract_bwts_number)
            
            # Create component analysis DataFrame
            component_data = []
            
            # Check for each component in each unit
            unique_units = sorted(bwts_data['Unit Number'].unique())
            
            for unit in unique_units:
                unit_name = f"BWTS #{unit}" if unit != 0 else "BWTS General"
                unit_data = bwts_data[bwts_data['Unit Number'] == unit]
                
                for component in self.components:
                    # Check if component appears in Sub Component Location or Title
                    has_component = False
                    if 'Sub Component Location' in unit_data.columns:
                        has_component = unit_data['Sub Component Location'].str.contains(component, case=False, na=False).any()
                    
                    if not has_component and 'Title' in unit_data.columns:
                        has_component = unit_data['Title'].str.contains(component, case=False, na=False).any()
                    
                    status = "Present" if has_component else "Missing"
                    component_data.append({
                        'Unit': unit_name,
                        'Component': component,
                        'Status': status
                    })
            
            # Convert to DataFrame
            component_df = pd.DataFrame(component_data)
            
            # Pivot to create a unit vs component matrix
            if not component_df.empty:
                pivot_df = component_df.pivot(index='Component', columns='Unit', values='Status')
                pivot_df = pivot_df.fillna("Missing")
                
                # Reset index to convert to regular DataFrame
                pivot_df = pivot_df.reset_index()
                
                return pivot_df
            else:
                return pd.DataFrame(columns=['Component', 'Status', 'Unit'])
                
        except Exception as e:
            print(f"Error analyzing BWTS components: {str(e)}")
            return pd.DataFrame(columns=['Component', 'Status', 'Unit'])
    
    def create_task_count_table(self, data):
        """Create task count analysis table for BWTS."""
        try:
            # Filter data for BWTS with more flexible patterns
            bwts_patterns = ['Ballast Water Treatment Plant', 'BWTS', 'Ballast Treatment']
            mask = data['Machinery Location'].str.contains('|'.join(bwts_patterns), case=False, na=False)
            bwts_data = data[mask].copy()
            print(f"Found {len(bwts_data)} BWTS records in create_task_count_table using patterns: {bwts_patterns}")
            
            if bwts_data.empty:
                return pd.DataFrame()
            
            # Count job codes by machinery location
            if 'Machinery Location' in bwts_data.columns and 'Job Code' in bwts_data.columns:
                task_counts = bwts_data.groupby('Machinery Location')['Job Code'].count().reset_index()
                task_counts.columns = ['Machinery Location', 'Task Count']
                
                # Sort by task count in descending order
                task_counts = task_counts.sort_values(by='Task Count', ascending=False)
                
                return task_counts
            else:
                print("Required columns not found for BWTS task count analysis.")
                return pd.DataFrame()
                
        except Exception as e:
            print(f"Error creating BWTS task count table: {str(e)}")
            return pd.DataFrame()
    
    def create_component_distribution(self, data):
        """Create component distribution analysis table."""
        try:
            # Filter data for BWTS with more flexible patterns
            bwts_patterns = ['Ballast Water Treatment Plant', 'BWTS', 'Ballast Treatment']
            mask = data['Machinery Location'].str.contains('|'.join(bwts_patterns), case=False, na=False)
            bwts_data = data[mask].copy()
            print(f"Found {len(bwts_data)} BWTS records in create_component_distribution using patterns: {bwts_patterns}")
            
            if bwts_data.empty:
                return pd.DataFrame()
                
            # Count occurrences of components
            component_counts = []
            
            for component in self.components:
                count = 0
                
                # Check Sub Component Location column
                if 'Sub Component Location' in bwts_data.columns:
                    count += bwts_data['Sub Component Location'].str.contains(component, case=False, na=False).sum()
                
                # Check Title column
                if 'Title' in bwts_data.columns:
                    count += bwts_data['Title'].str.contains(component, case=False, na=False).sum()
                
                component_counts.append({
                    'Component': component,
                    'Count': count
                })
            
            # Convert to DataFrame
            component_df = pd.DataFrame(component_counts)
            
            # Sort by count in descending order
            component_df = component_df.sort_values(by='Count', ascending=False)
            
            return component_df
            
        except Exception as e:
            print(f"Error creating BWTS component distribution: {str(e)}")
            return pd.DataFrame(columns=['Component', 'Count'])
    
    def process_reference_data(self, data, ref_sheet, preferred_sheet=None):
        """Process reference data for BWTS using the approach from the code snippet.
        
        Args:
            data: DataFrame containing the machinery data
            ref_sheet: Path to the reference Excel file
            preferred_sheet: Optional specific sheet name to use for BWTS model
        """
        try:
            # Create copies to avoid modifying original data
            data_copy = data.copy()
            
            # Filter data for BWTS with more flexible patterns
            bwts_patterns = ['Ballast Water Treatment Plant', 'BWTS', 'Ballast Treatment']
            mask = data_copy['Machinery Location'].str.contains('|'.join(bwts_patterns), case=False, na=False)
            filtered_dfBWTSjobs = data_copy[mask].copy()
            print(f"Found {len(filtered_dfBWTSjobs)} BWTS records in process_reference_data using patterns: {bwts_patterns}")
            
            # Read the reference sheet using the uploaded sheet path
            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            
            # First priority: Use the preferred sheet if specified and it exists
            bwts_sheet = None
            if preferred_sheet is not None and preferred_sheet in ref_sheet_names:
                bwts_sheet = preferred_sheet
                print(f"Using preferred reference sheet for BWTS model: {bwts_sheet}")
            # Second priority: Look for 'BWTS' sheet
            elif 'BWTS' in ref_sheet_names:
                bwts_sheet = 'BWTS'
            
            # Third priority: Try to find any sheet with 'BWTS' or 'Ballast' in the name
            if bwts_sheet is None:
                for sheet in ref_sheet_names:
                    if 'BWTS' in sheet or 'Ballast' in sheet.lower():
                        bwts_sheet = sheet
                        break
            
            # Last resort: If still no BWTS-specific sheet found, use the first sheet
            if bwts_sheet is None:
                bwts_sheet = ref_sheet_names[0]
                print(f"No BWTS sheet found, using the first sheet: {bwts_sheet}")
            else:
                print(f"Using reference sheet: {bwts_sheet}")
            
            # Read the reference sheet
            dfBWTS = pd.read_excel(ref_sheet, sheet_name=bwts_sheet)
            
            # Skip further processing if no data
            if filtered_dfBWTSjobs.empty or dfBWTS.empty:
                return pd.DataFrame({
                    'Job Code': ['No data found'],
                    'Title': ['No matching data between current and reference'],
                    'Frequency': ['N/A']
                })
            
            # Ensure all Job Code columns are strings
            if 'Job Code' in filtered_dfBWTSjobs.columns:
                filtered_dfBWTSjobs['Job Codecopy'] = filtered_dfBWTSjobs['Job Code'].astype(str)
            else:
                print("Job Code column not found in BWTS data")
                return pd.DataFrame({
                    'Job Code': ['Column Error'],
                    'Title': ['Job Code column not found in data'],
                    'Frequency': ['N/A']
                })
            
            # Check which column to use for job code in reference data
            job_code_col = None
            for possible_col in ['UI Job Code', 'Job Code', 'JobCode', 'Code']:
                if possible_col in dfBWTS.columns:
                    job_code_col = possible_col
                    break
            
            if job_code_col is None:
                print("No job code column found in reference data")
                # Create a sample of the columns we have
                col_sample = ", ".join(dfBWTS.columns.tolist()[:5]) + "..."
                return pd.DataFrame({
                    'Job Code': ['Column Error'],
                    'Title': [f'No job code column found in reference. Available columns: {col_sample}'],
                    'Frequency': ['N/A']
                })
            
            # Find job codes in dfBWTS that are not in filtered_dfBWTSjobs
            dfBWTS['Job Code'] = dfBWTS[job_code_col].astype(str)
            filtered_dfBWTSjobs['Job Codecopy'] = filtered_dfBWTSjobs['Job Codecopy'].astype(str)
            
            print(f"Sample reference job codes: {dfBWTS['Job Code'].iloc[:5].tolist()}")
            print(f"Sample current job codes: {filtered_dfBWTSjobs['Job Codecopy'].iloc[:5].tolist()}")
            print(f"Total reference jobs: {len(dfBWTS)}")
            print(f"Total current BWTS jobs: {len(filtered_dfBWTSjobs)}")
            
            missingjobsbwtsresult = dfBWTS[~dfBWTS['Job Code'].isin(filtered_dfBWTSjobs['Job Codecopy'])]
            print(f"Total missing jobs found: {len(missingjobsbwtsresult)}")
            
            # Drop Remarks column if it exists
            if 'Remarks' in missingjobsbwtsresult.columns:
                missingjobsbwtsresult = missingjobsbwtsresult.drop(columns=['Remarks'])
                
            # Reset the index of the combined DataFrame
            missingjobsbwtsresult.reset_index(drop=True, inplace=True)
            
            return missingjobsbwtsresult
                
        except Exception as e:
            print(f"Error processing BWTS reference data: {str(e)}")
            # Create an error row for display
            error_row = pd.DataFrame({
                'Job Code': ['Error'],
                'Title': [f'Unable to process reference data: {str(e)}'],
                'Frequency': ['N/A']
            })
            return error_row