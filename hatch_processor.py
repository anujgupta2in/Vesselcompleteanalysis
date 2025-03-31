import pandas as pd
import numpy as np
import re

class HatchProcessor:
    def __init__(self):
        """Initialize HatchProcessor with component list."""
        self.components = [
            'Hatch Cover',
            'Hatch Coaming',
            'Hatch Hydraulic System',
            'Hatch Cylinders',
            'Hatch Cleats',
            'Hatch Sealing',
            'Hatch Panel',
            'Hatch Control System',
            'Hatch Motor',
            'Hatch Safety System'
        ]
    
    def extract_running_hours(self, data):
        """Extract running hours for hatches. Note: Hatches typically don't have running hours,
        but method is included for consistency with other processors."""
        try:
            # Filter data for Hatches with flexible patterns
            hatch_patterns = ['Hatch', 'Cargo Hatch', 'Cargo Opening']
            mask = data['Machinery Location'].str.contains('|'.join(hatch_patterns), case=False, na=False)
            hatch_data = data[mask].copy()
            print(f"Found {len(hatch_data)} Hatch records using patterns: {hatch_patterns}")
            
            if hatch_data.empty:
                # Return empty DataFrame with expected columns to avoid errors
                return pd.DataFrame(columns=['Hatch Unit', 'Running Hours'])
            
            # Initialize empty list to store running hours data
            running_hours_data = []
            
            # Get unique Hatch units
            def extract_hatch_number(location):
                if not isinstance(location, str):
                    return "Unknown Hatch Unit"
                    
                # Try multiple patterns to extract unit numbers
                hatch_patterns = [
                    r'Hatch#?(\d+)',
                    r'Cargo Hatch#?(\d+)',
                    r'Cargo Opening#?(\d+)'
                ]
                
                for pattern in hatch_patterns:
                    match = re.search(pattern, location, re.IGNORECASE)
                    if match:
                        return f"Hatch #{match.group(1)}"
                
                # If no match, just return the location as is (with a safety check)
                return location[:100] if len(location) > 100 else location
            
            # Make sure there's a Hatch Unit column
            hatch_data['Hatch Unit'] = hatch_data['Machinery Location'].apply(extract_hatch_number)
            unique_units = hatch_data['Hatch Unit'].unique()
            
            # For each unit, get the running hours (typically 0 for hatches as they don't have running hours)
            for unit in unique_units:
                unit_data = hatch_data[hatch_data['Hatch Unit'] == unit]
                
                # Get the running hours for this unit if they exist
                running_hours = 0  # Default to 0
                if 'Machinery Running Hours' in unit_data.columns:
                    # Convert to numeric, coercing errors to NaN
                    unit_data['Machinery Running Hours'] = pd.to_numeric(unit_data['Machinery Running Hours'], errors='coerce')
                    running_hours_values = unit_data['Machinery Running Hours'].dropna()
                    if not running_hours_values.empty:
                        running_hours = running_hours_values.iloc[0]
                
                # Add to our list with number format to avoid serialization issues
                running_hours_data.append({
                    'Hatch Unit': str(unit),  # Ensure string type
                    'Running Hours': float(running_hours) if pd.notna(running_hours) else 0.0  # Ensure numeric type
                })
            
            # Convert to DataFrame and ensure proper types
            result_df = pd.DataFrame(running_hours_data)
            if result_df.empty:
                return pd.DataFrame(columns=['Hatch Unit', 'Running Hours'])
                
            return result_df
            
        except Exception as e:
            print(f"Error extracting Hatch running hours: {str(e)}")
            # Return empty DataFrame with expected columns
            return pd.DataFrame(columns=['Hatch Unit', 'Running Hours'])
    
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
            print(f"Error formatting Hatch unit data: {str(e)}")
            return unit_data
    
    def process_job_code(self, data, job_codes, job_description, unit_pattern=r'(?:Hatch|Cargo Hatch|Cargo Opening)#?(\d+)'):
        """Process job codes for hatches."""
        try:
            # Filter data for Hatches with flexible patterns
            hatch_patterns = ['Hatch', 'Cargo Hatch', 'Cargo Opening']
            mask = data['Machinery Location'].str.contains('|'.join(hatch_patterns), case=False, na=False)
            hatch_data = data[mask].copy()
            print(f"Found {len(hatch_data)} Hatch records in process_job_code using patterns: {hatch_patterns}")
            
            if hatch_data.empty:
                return pd.DataFrame()
                
            # Process the data
            hatch_data['Unit'] = hatch_data['Machinery Location'].str.extract(unit_pattern, expand=False)
            hatch_data['Unit'] = hatch_data['Unit'].fillna('General')
            hatch_data['Unit'] = 'Unit #' + hatch_data['Unit']
            
            # Create pivot table
            if 'Job Code' in hatch_data.columns and job_codes in hatch_data.columns and job_description in hatch_data.columns:
                pivot_df = pd.pivot_table(
                    hatch_data,
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
                print(f"Required columns not found for Hatch job code processing.")
                return pd.DataFrame()
                
        except Exception as e:
            print(f"Error processing Hatch job codes: {str(e)}")
            return pd.DataFrame()
    
    def get_maintenance_data(self, data):
        """Get Hatch maintenance data."""
        try:
            # Filter data for Hatches with flexible patterns
            hatch_patterns = ['Hatch', 'Cargo Hatch', 'Cargo Opening']
            mask = data['Machinery Location'].str.contains('|'.join(hatch_patterns), case=False, na=False)
            hatch_data = data[mask].copy()
            print(f"Found {len(hatch_data)} Hatch records in get_maintenance_data using patterns: {hatch_patterns}")
            
            if hatch_data.empty:
                return pd.DataFrame()
            
            # Select relevant columns
            if 'Job Code' in hatch_data.columns and 'Title' in hatch_data.columns and 'Calculated Due Date' in hatch_data.columns:
                maintenance_data = hatch_data[['Job Code', 'Title', 'Calculated Due Date', 'Machinery Location', 'Job Status']].copy()
                
                # Sort by due date
                if 'Calculated Due Date' in maintenance_data.columns:
                    maintenance_data = maintenance_data.sort_values(by='Calculated Due Date')
                
                return maintenance_data
            else:
                print("Required columns not found for Hatch maintenance data.")
                return pd.DataFrame()
                
        except Exception as e:
            print(f"Error getting Hatch maintenance data: {str(e)}")
            return pd.DataFrame()
    
    def analyze_components(self, data):
        """Analyze component presence for Hatches."""
        try:
            # Filter data for Hatches with flexible patterns
            hatch_patterns = ['Hatch', 'Cargo Hatch', 'Cargo Opening']
            mask = data['Machinery Location'].str.contains('|'.join(hatch_patterns), case=False, na=False)
            hatch_data = data[mask].copy()
            print(f"Found {len(hatch_data)} Hatch records in analyze_components using patterns: {hatch_patterns}")
            
            if hatch_data.empty:
                return pd.DataFrame(), 0
                
            # Extract Hatch unit numbers
            def extract_hatch_number(location):
                # Try multiple patterns to extract unit numbers
                hatch_patterns = [
                    r'Hatch#?(\d+)',
                    r'Cargo Hatch#?(\d+)',
                    r'Cargo Opening#?(\d+)'
                ]
                
                for pattern in hatch_patterns:
                    match = re.search(pattern, location, re.IGNORECASE)
                    if match:
                        return int(match.group(1))
                
                return 0  # default value for general items
                
            hatch_data['Unit Number'] = hatch_data['Machinery Location'].apply(extract_hatch_number)
            
            # Create component analysis DataFrame
            component_data = []
            
            # Check for each component in each unit
            unique_units = sorted(hatch_data['Unit Number'].unique())
            
            for unit in unique_units:
                unit_name = f"Hatch #{unit}" if unit != 0 else "Hatch General"
                unit_data = hatch_data[hatch_data['Unit Number'] == unit]
                
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
            
            # Count missing components
            missing_count = (component_df['Status'] == 'Missing').sum()
            
            # Format for display
            if not component_df.empty:
                return component_df, missing_count
            else:
                return pd.DataFrame(columns=['Component', 'Status', 'Unit']), 0
                
        except Exception as e:
            print(f"Error analyzing Hatch components: {str(e)}")
            return pd.DataFrame(columns=['Component', 'Status', 'Unit']), 0
    
    def create_task_count_table(self, data):
        """Create task count analysis table for Hatches."""
        try:
            # Filter data for Hatches with flexible patterns
            hatch_patterns = ['Hatch', 'Cargo Hatch', 'Cargo Opening']
            mask = data['Machinery Location'].str.contains('|'.join(hatch_patterns), case=False, na=False)
            hatch_data = data[mask].copy()
            print(f"Found {len(hatch_data)} Hatch records in create_task_count_table using patterns: {hatch_patterns}")
            
            if hatch_data.empty:
                return pd.DataFrame()
            
            # Count job codes by machinery location
            if 'Machinery Location' in hatch_data.columns and 'Job Code' in hatch_data.columns:
                task_counts = hatch_data.groupby('Machinery Location')['Job Code'].count().reset_index()
                task_counts.columns = ['Machinery Location', 'Task Count']
                
                # Sort by task count in descending order
                task_counts = task_counts.sort_values(by='Task Count', ascending=False)
                
                return task_counts
            else:
                print("Required columns not found for Hatch task count analysis.")
                return pd.DataFrame()
                
        except Exception as e:
            print(f"Error creating Hatch task count table: {str(e)}")
            return pd.DataFrame()
    
    def create_component_distribution(self, data):
        """Create component distribution analysis table."""
        try:
            # Filter data for Hatches with flexible patterns
            hatch_patterns = ['Hatch', 'Cargo Hatch', 'Cargo Opening']
            mask = data['Machinery Location'].str.contains('|'.join(hatch_patterns), case=False, na=False)
            hatch_data = data[mask].copy()
            print(f"Found {len(hatch_data)} Hatch records in create_component_distribution using patterns: {hatch_patterns}")
            
            if hatch_data.empty:
                return pd.DataFrame()
                
            # Count occurrences of components
            component_counts = []
            
            for component in self.components:
                count = 0
                
                # Check Sub Component Location column
                if 'Sub Component Location' in hatch_data.columns:
                    count += hatch_data['Sub Component Location'].str.contains(component, case=False, na=False).sum()
                
                # Check Title column
                if 'Title' in hatch_data.columns:
                    count += hatch_data['Title'].str.contains(component, case=False, na=False).sum()
                
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
            print(f"Error creating Hatch component distribution: {str(e)}")
            return pd.DataFrame(columns=['Component', 'Count'])
    
    def process_reference_data(self, data, ref_sheet, preferred_sheet='Hatch'):
        """Process reference data for Hatches.
        
        Args:
            data: DataFrame containing the machinery data
            ref_sheet: Path to the reference Excel file
            preferred_sheet: Optional specific sheet name to use for Hatches
            
        Returns:
            DataFrame with reference jobs for hatch covers
        """
        try:
            # Create copies to avoid modifying original data
            data_copy = data.copy()
            
            # Filter data for Hatches with flexible patterns
            hatch_patterns = ['Hatch', 'Cargo Hatch', 'Cargo Opening']
            mask = data_copy['Machinery Location'].str.contains('|'.join(hatch_patterns), case=False, na=False)
            filtered_dfHatchjobs = data_copy[mask].copy()
            print(f"Found {len(filtered_dfHatchjobs)} Hatch records in process_reference_data using patterns: {hatch_patterns}")
            
            # Read the reference sheet using the uploaded sheet path
            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            
            # First priority: Use the preferred sheet if specified and it exists
            hatch_sheet = None
            if preferred_sheet is not None and preferred_sheet in ref_sheet_names:
                hatch_sheet = preferred_sheet
                print(f"Using preferred reference sheet for Hatch analysis: {hatch_sheet}")
            # Second priority: Look for 'Hatch' sheet
            elif 'Hatch' in ref_sheet_names:
                hatch_sheet = 'Hatch'
            
            # Third priority: Try to find any sheet with 'Hatch' or 'Cargo' in the name
            if hatch_sheet is None:
                for sheet in ref_sheet_names:
                    if 'Hatch' in sheet or 'Cargo' in sheet:
                        hatch_sheet = sheet
                        break
            
            # Last resort: If still no Hatch-specific sheet found, use the first sheet
            if hatch_sheet is None:
                hatch_sheet = ref_sheet_names[0]
                print(f"No Hatch sheet found, using the first sheet: {hatch_sheet}")
            else:
                print(f"Using reference sheet: {hatch_sheet}")
            
            # Read the reference sheet
            dfHatch = pd.read_excel(ref_sheet, sheet_name=hatch_sheet)
            
            # Skip further processing if no data
            if filtered_dfHatchjobs.empty or dfHatch.empty:
                return pd.DataFrame({
                    'Job Code': ['No data found'],
                    'Title': ['No matching data between current and reference'],
                    'Source': ['N/A']
                })
            
            # Ensure all Job Code columns are strings
            if 'Job Code' in filtered_dfHatchjobs.columns:
                filtered_dfHatchjobs['Job Codecopy'] = filtered_dfHatchjobs['Job Code'].astype(str)
            else:
                print("Job Code column not found in Hatch data")
                return pd.DataFrame({
                    'Job Code': ['Column Error'],
                    'Title': ['Job Code column not found in data'],
                    'Source': ['N/A']
                })
            
            # Check which column to use for job code in reference data
            job_code_col = None
            for possible_col in ['UI Job Code', 'Job Code', 'JobCode', 'Code']:
                if possible_col in dfHatch.columns:
                    job_code_col = possible_col
                    break
            
            if job_code_col is None:
                print("No job code column found in reference data")
                # Create a sample of the columns we have
                col_sample = ", ".join(dfHatch.columns.tolist()[:5]) + "..."
                return pd.DataFrame({
                    'Job Code': ['Column Error'],
                    'Title': [f'No job code column found in reference. Available columns: {col_sample}'],
                    'Source': ['N/A']
                })
            
            # Find title column in reference data
            title_col = None
            for possible_col in ['J3 Job Title', 'Job Title', 'Title', 'Description']:
                if possible_col in dfHatch.columns:
                    title_col = possible_col
                    break
                    
            if title_col is None:
                title_col = dfHatch.columns[1] if len(dfHatch.columns) > 1 else 'No Title'
            
            # Find job codes in dfHatch that are not in filtered_dfHatchjobs
            dfHatch['Job Code'] = dfHatch[job_code_col].astype(str)
            filtered_dfHatchjobs['Job Codecopy'] = filtered_dfHatchjobs['Job Codecopy'].astype(str)
            
            print(f"Sample reference job codes: {dfHatch['Job Code'].iloc[:5].tolist()}")
            print(f"Sample current job codes: {filtered_dfHatchjobs['Job Codecopy'].iloc[:5].tolist()}")
            print(f"Total reference jobs: {len(dfHatch)}")
            print(f"Total current Hatch jobs: {len(filtered_dfHatchjobs)}")
            
            # Create a DataFrame for Reference Jobs with missing jobs
            missing_jobs = dfHatch[~dfHatch['Job Code'].isin(filtered_dfHatchjobs['Job Codecopy'])].copy()
            
            # # Add a status column to show these are missing jobs
            missing_jobs['Status'] = 'Missing'
            
            # Get matching jobs as well (jobs that exist in both reference and current data)
            matching_jobs = dfHatch[dfHatch['Job Code'].isin(filtered_dfHatchjobs['Job Codecopy'])].copy()
            matching_jobs['Status'] = 'Present'
            
            # Combine both missing and matching jobs
            reference_jobs = pd.concat([missing_jobs,matching_jobs], ignore_index=True)
            reference_jobs['Source'] = 'Reference Sheet'
            
            # Create a clean DataFrame with selected columns
            selected_columns = ['Job Code']
            if title_col in reference_jobs.columns:
                selected_columns.append(title_col)
            selected_columns.extend(['Status', 'Source'])
            
            result_df = reference_jobs[selected_columns].copy()
            
            # Rename columns for consistency
            column_mapping = {}
            if title_col != 'Title' and title_col in result_df.columns:
                column_mapping[title_col] = 'Title'
                
            if column_mapping:
                result_df = result_df.rename(columns=column_mapping)
            
            # Sort jobs by status (Missing first) and then by job code
            result_df = result_df.sort_values(by=['Status', 'Job Code'])
            
            # Reset the index of the DataFrame
            result_df.reset_index(drop=True, inplace=True)
            
            print(f"Total reference jobs returned: {len(result_df)}")
            return missing_jobs
                
        except Exception as e:
            print(f"Error processing Hatch reference data: {str(e)}")
            # Create an error row for display
            error_row = pd.DataFrame({
                'Job Code': ['Error'],
                'Title': [f'Unable to process reference data: {str(e)}'],
                'Source': ['N/A']
            })
            return error_row
            
    def create_reference_pivot_table(self, data, ref_sheet, preferred_sheet='Hatch'):
        """Create a pivot table from reference data for Hatches, similar to Purifier implementation.
        
        Args:
            data: DataFrame containing the machinery data
            ref_sheet: Path to the reference Excel file
            preferred_sheet: Optional specific sheet name to use for Hatches
            
        Returns:
            Pivot table DataFrame showing job distribution across machinery locations
        """
        try:
            # Create copies to avoid modifying original data
            data_copy = data.copy()
            
            # Filter data for Hatches with flexible patterns
            hatch_patterns = ['Hatch', 'Cargo Hatch', 'Cargo Opening']
            mask = data_copy['Machinery Location'].str.contains('|'.join(hatch_patterns), case=False, na=False)
            filtered_dfHatchjobs = data_copy[mask].copy()
            print(f"Found {len(filtered_dfHatchjobs)} Hatch records in create_reference_pivot_table using patterns: {hatch_patterns}")
            
            # Skip if no data
            if filtered_dfHatchjobs.empty:
                return pd.DataFrame()
                
            # Read the reference sheet using the uploaded sheet path
            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            
            # First priority: Use the preferred sheet if specified and it exists
            hatch_sheet = None
            if preferred_sheet is not None and preferred_sheet in ref_sheet_names:
                hatch_sheet = preferred_sheet
            # Second priority: Look for 'Hatch' sheet
            elif 'Hatch' in ref_sheet_names:
                hatch_sheet = 'Hatch'
            
            # Third priority: Try to find any sheet with 'Hatch' or 'Cargo' in the name
            if hatch_sheet is None:
                for sheet in ref_sheet_names:
                    if 'Hatch' in sheet or 'Cargo' in sheet:
                        hatch_sheet = sheet
                        break
            
            # Last resort: If still no Hatch-specific sheet found, use the first sheet
            if hatch_sheet is None:
                hatch_sheet = ref_sheet_names[0]
                print(f"No Hatch sheet found for pivot, using the first sheet: {hatch_sheet}")
            else:
                print(f"Using reference sheet for pivot: {hatch_sheet}")
            
            # Read the reference sheet
            dfHatch = pd.read_excel(ref_sheet, sheet_name=hatch_sheet)
            
            # Ensure all Job Code columns are strings
            if 'Job Code' in filtered_dfHatchjobs.columns:
                filtered_dfHatchjobs['Job Codecopy'] = filtered_dfHatchjobs['Job Code'].astype(str)
            else:
                print("Job Code column not found in Hatch data for pivot")
                return pd.DataFrame()
            
            # Check which column to use for job code in reference data
            job_code_col = None
            for possible_col in ['UI Job Code', 'Job Code', 'JobCode', 'Code']:
                if possible_col in dfHatch.columns:
                    job_code_col = possible_col
                    break
            
            if job_code_col is None:
                print("No job code column found in reference data for pivot")
                return pd.DataFrame()
            
            # Find title column in reference data
            title_col = None
            for possible_col in ['J3 Job Title', 'Job Title', 'Title', 'Description']:
                if possible_col in dfHatch.columns:
                    title_col = possible_col
                    break
                    
            if title_col is None:
                title_col = dfHatch.columns[1] if len(dfHatch.columns) > 1 else 'No Title'
            
            # Convert job code columns to string type
            dfHatch[job_code_col] = dfHatch[job_code_col].astype(str)
            filtered_dfHatchjobs['Job Codecopy'] = filtered_dfHatchjobs['Job Codecopy'].astype(str)
            
            # Merge filtered_dfHatchjobs with dfHatch on matching job codes
            result_dfHatchJobs = filtered_dfHatchjobs.merge(
                dfHatch, 
                left_on='Job Codecopy', 
                right_on=job_code_col, 
                suffixes=('_filtered', '_ref')
            )
            
            # Reset index of the result DataFrame
            result_dfHatchJobs.reset_index(drop=True, inplace=True)
            
            # Create pivot table if we have the necessary columns
            if title_col is not None and 'Machinery Location' in result_dfHatchJobs.columns:
                pivot_table_resultHatchJobs = result_dfHatchJobs.pivot_table(
                    index=title_col, 
                    columns='Machinery Location', 
                    values='Job Codecopy', 
                    aggfunc='count'
                )
                return pivot_table_resultHatchJobs
            else:
                print("Required columns not found for Hatch pivot table")
                return pd.DataFrame()
                
        except Exception as e:
            print(f"Error creating Hatch reference pivot table: {str(e)}")
            return pd.DataFrame()