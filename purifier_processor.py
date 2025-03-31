import pandas as pd
import numpy as np
import re

class PurifierProcessor:
    def __init__(self):
        """Initialize PurifierProcessor with component list."""
        self.components = [
            'Bowl - PU',
            'Gear - PU',
            'Clutch/Brake - PU',
            'Motor - PU',
            'Separator - PU',
            'Control System - PU',
            'Oil Supply Line - PU',
            'Lubrication System - PU',
            'Seal - PU',
            'Frame/Foundation - PU',
            'Drive Assembly - PU',
            'Heater - PU'
        ]
        
        # Define the pattern to extract purifier numbers
        self.purifier_pattern = r'Purifier.*?#?(\d+)'
        
    def extract_running_hours(self, data):
        """Extract running hours for purifiers."""
        try:
            if 'Machinery Location' not in data.columns or 'Running Hours' not in data.columns:
                return pd.DataFrame(columns=['Purifier', 'Running Hours'])
            
            # Create a copy to avoid SettingWithCopyWarning
            purifier_data = data.copy()
            
            # Filter for purifier data
            purifier_data = purifier_data[purifier_data['Machinery Location'].str.contains('Purifier', case=False, na=False)]
            
            # If no purifier data, return empty DataFrame with correct columns
            if purifier_data.empty:
                return pd.DataFrame(columns=['Purifier', 'Running Hours'])
            
            # Extract purifier numbers safely
            def extract_purifier_number(location):
                location_str = str(location)
                match = re.search(self.purifier_pattern, location_str)
                if match and match.groups():
                    return match.group(1)
                else:
                    return 'Unknown'
            
            purifier_data['Purifier'] = purifier_data['Machinery Location'].apply(extract_purifier_number)
            
            # Group by purifier and take the max running hours
            running_hours = purifier_data.groupby('Purifier')['Running Hours'].max().reset_index()
            
            # Format the running hours
            running_hours['Running Hours'] = running_hours['Running Hours'].apply(
                lambda x: f"{int(x):,}" if pd.notnull(x) and x != 0 else "N/A"
            )
            
            # If still empty, create sample data
            if running_hours.empty:
                return pd.DataFrame({
                    'Purifier': ['No purifiers found'],
                    'Running Hours': ['N/A']
                })
                
            return running_hours
            
        except Exception as e:
            print(f"Error extracting purifier running hours: {str(e)}")
            # Return visible error message
            return pd.DataFrame({
                'Purifier': ['Error'],
                'Running Hours': [f'Error: {str(e)}']
            })
            
    def format_unit_data(self, unit_data):
        """Format unit data for display."""
        try:
            # Format dictionary for display
            formatted_data = []
            for unit, tasks in unit_data.items():
                # Format each unit
                if isinstance(tasks, list):
                    task_text = "<br>".join(tasks)
                else:
                    task_text = str(tasks)
                
                formatted_data.append({
                    "Purifier": unit,
                    "Maintenance Tasks": task_text
                })
            
            # Convert to DataFrame for display
            return pd.DataFrame(formatted_data)
        
        except Exception as e:
            print(f"Error formatting purifier data: {str(e)}")
            return pd.DataFrame(columns=['Purifier', 'Maintenance Tasks'])
            
    def process_job_code(self, data, job_codes, job_description, unit_pattern=r'Purifier.*?#?(\d+)'):
        """Process job codes for purifiers."""
        try:
            # Make a copy of the data to avoid modifying the original
            data_copy = data.copy()
            
            # Ensure Job Code is string type to use string methods
            data_copy['Job Code'] = data_copy['Job Code'].astype(str)
            
            # Filter the data for purifier job codes
            purifier_jobs = data_copy[data_copy['Machinery Location'].str.contains('Purifier', case=False, na=False)].copy()
            
            # Initialize dictionary to hold tasks by unit
            unit_tasks = {}
            
            # Process each job code
            for job_code in job_codes:
                # Filter for this job code
                matched_jobs = purifier_jobs[purifier_jobs['Job Code'].str.contains(job_code, case=False)]
                
                for _, row in matched_jobs.iterrows():
                    # Get the machine name and extract unit number
                    machine_name = row['Machinery Location']
                    unit_match = re.search(unit_pattern, str(machine_name))
                    
                    if unit_match:
                        unit_number = unit_match.group(1)
                        unit_id = f"Purifier #{unit_number}"
                        
                        # Get the task description
                        task = row.get(job_description, 'No description')
                        # Convert task to string if it's not already
                        task = str(task) if not pd.isna(task) else 'No description'
                        
                        # Add to unit tasks
                        if unit_id not in unit_tasks:
                            unit_tasks[unit_id] = []
                        
                        # Add task if not already present
                        if task not in unit_tasks[unit_id]:
                            unit_tasks[unit_id].append(task)
            
            # If no tasks were found, create a sample for display
            if not unit_tasks:
                unit_tasks["Purifier Data"] = ["No purifier-specific job codes found. Please check your data."]
                
            return unit_tasks
            
        except Exception as e:
            print(f"Error processing purifier job codes: {str(e)}")
            # Return a simple dict with error message for visibility
            return {"Error": [f"Could not process job codes: {str(e)}"]}
            
    def get_maintenance_data(self, data):
        """Get purifier maintenance data."""
        try:
            # Define key job codes for purifier maintenance
            job_codes = ['PU-', 'PURIF', 'SEPAR']
            
            # Process job codes
            maintenance_data = self.process_job_code(
                data,
                job_codes,
                'Task Description',
                unit_pattern=self.purifier_pattern
            )
            
            # Format for display
            return self.format_unit_data(maintenance_data)
            
        except Exception as e:
            print(f"Error getting purifier maintenance data: {str(e)}")
            return pd.DataFrame(columns=['Purifier', 'Maintenance Tasks'])
            
    def analyze_components(self, data):
        """Analyze component presence for purifiers."""
        try:
            # Create a copy of the data
            purifier_data = data[data['Machinery Location'].str.contains('Purifier', case=False, na=False)].copy()
            
            # If no purifier data, return empty DataFrame
            if purifier_data.empty:
                return pd.DataFrame(columns=['Purifier'] + [comp.split(' - ')[0] for comp in self.components])
            
            # Extract purifier numbers safely
            def extract_purifier_number(location):
                location_str = str(location)
                match = re.search(self.purifier_pattern, location_str)
                if match and match.groups():
                    return match.group(1)
                else:
                    return 'Unknown'
            
            purifier_data['Purifier'] = purifier_data['Machinery Location'].apply(extract_purifier_number)
            
            # Check component presence based on Sub Component Location
            result = []
            
            for purifier_num in sorted(purifier_data['Purifier'].unique()):
                purifier_id = f"Purifier #{purifier_num}"
                purifier_components = purifier_data[purifier_data['Purifier'] == purifier_num]
                
                component_status = {}
                
                # Check each component
                for component in self.components:
                    # Check if component is mentioned in Sub Component Location
                    if 'Sub Component Location' in purifier_components.columns:
                        has_component = purifier_components['Sub Component Location'].str.contains(
                            component, case=False, na=False
                        ).any()
                    else:
                        has_component = False
                    
                    component_name = component.split(' - ')[0]  # Remove the "- PU" suffix for display
                    component_status[component_name] = "✓" if has_component else "✗"
                
                # Add to result
                result.append({"Purifier": purifier_id, **component_status})
            
            # Return the DataFrame
            if result:
                return pd.DataFrame(result)
            else:
                # Create a sample row if no components found
                sample_row = {"Purifier": "No purifiers found"}
                for component in self.components:
                    component_name = component.split(' - ')[0]
                    sample_row[component_name] = "✗"
                return pd.DataFrame([sample_row])
            
        except Exception as e:
            print(f"Error analyzing purifier components: {str(e)}")
            # Return a DataFrame with error message
            return pd.DataFrame({
                'Purifier': ['Error'],
                'Error Message': [f'Error analyzing components: {str(e)}']
            })
            
    def create_task_count_table(self, data):
        """Create task count analysis table for purifiers."""
        try:
            # Filter for purifier data
            purifier_data = data[data['Machinery Location'].str.contains('Purifier', case=False, na=False)].copy()
            
            # If no purifier data, return empty DataFrame
            if purifier_data.empty:
                return pd.DataFrame()
            
            # Count tasks by machinery location
            task_counts = purifier_data.groupby('Machinery Location').size().reset_index(name='Task Count')
            
            # Sort by task count in descending order
            task_counts = task_counts.sort_values('Task Count', ascending=False)
            
            # Clean up machinery location for display
            task_counts['Machinery Location'] = task_counts['Machinery Location'].apply(
                lambda x: str(x).replace('  ', ' ').strip()
            )
            
            return task_counts
            
        except Exception as e:
            print(f"Error creating purifier task count table: {str(e)}")
            return pd.DataFrame()
            
    def create_component_distribution(self, data):
        """Create component distribution analysis table."""
        try:
            # Filter for purifier data
            purifier_data = data[data['Machinery Location'].str.contains('Purifier', case=False, na=False)].copy()
            
            # Count component distribution
            if 'Sub Component Location' in purifier_data.columns:
                # Check for each component in our list
                component_counts = {}
                
                for component in self.components:
                    # Count occurrences of component in Sub Component Location
                    component_name = component.split(' - ')[0]  # Remove the "- PU" suffix for display
                    component_counts[component_name] = purifier_data['Sub Component Location'].str.contains(
                        component, case=False, na=False
                    ).sum()
                
                # Convert to DataFrame
                component_df = pd.DataFrame({
                    'Component': list(component_counts.keys()),
                    'Count': list(component_counts.values())
                })
                
                # Sort by count
                component_df = component_df.sort_values('Count', ascending=False)
                
                return component_df
            else:
                return pd.DataFrame(columns=['Component', 'Count'])
                
        except Exception as e:
            print(f"Error creating purifier component distribution: {str(e)}")
            return pd.DataFrame(columns=['Component', 'Count'])
            
    def process_reference_data(self, data, ref_sheet):
        """Process reference data for purifiers using the approach from the code snippet."""
        try:
            # Create copies to avoid modifying original data
            data_copy = data.copy()
            
            # Filter data for purifiers
            filtered_dfpurifierjobs = data_copy[data_copy['Machinery Location'].str.contains('Purifier', case=False, na=False)].copy()
            
            # Use the specified fixed reference sheet with sheet name 'Purifiers'
            # Read the reference sheet using the uploaded sheet path
            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            
            # Look for 'Purifiers' sheet specifically first
            purifier_sheet = 'Purifiers' if 'Purifiers' in ref_sheet_names else None
            
            # If not found, try to find any sheet with 'Purifier' or 'PU' in the name
            if purifier_sheet is None:
                for sheet in ref_sheet_names:
                    if 'Purifier' in sheet.lower() or 'PU' in sheet:
                        purifier_sheet = sheet
                        break
            
            # If still no purifier-specific sheet found, use the first sheet
            if purifier_sheet is None:
                purifier_sheet = ref_sheet_names[0]
                print(f"No purifier sheet found, using the first sheet: {purifier_sheet}")
            else:
                print(f"Using reference sheet: {purifier_sheet}")
            
            # Read the reference sheet
            dfpurifiers = pd.read_excel(ref_sheet, sheet_name=purifier_sheet)
            
            # Skip further processing if no data
            if filtered_dfpurifierjobs.empty or dfpurifiers.empty:
                return pd.DataFrame({
                    'Job Code': ['No data found'],
                    'Title': ['No matching data between current and reference'],
                    'Frequency': ['N/A']
                })
            
            # Ensure all Job Code columns are strings
            if 'Job Code' in filtered_dfpurifierjobs.columns:
                filtered_dfpurifierjobs['Job Codecopy'] = filtered_dfpurifierjobs['Job Code'].astype(str)
            else:
                print("Job Code column not found in purifier data")
                return pd.DataFrame({
                    'Job Code': ['Column Error'],
                    'Title': ['Job Code column not found in data'],
                    'Frequency': ['N/A']
                })
            
            # Check which column to use for job code in reference data
            job_code_col = None
            for possible_col in ['UI Job Code', 'Job Code', 'JobCode', 'Code']:
                if possible_col in dfpurifiers.columns:
                    job_code_col = possible_col
                    break
            
            if job_code_col is None:
                print("No job code column found in reference data")
                # Create a sample of the columns we have
                col_sample = ", ".join(dfpurifiers.columns.tolist()[:5]) + "..."
                return pd.DataFrame({
                    'Job Code': ['Column Error'],
                    'Title': [f'No job code column found in reference. Available columns: {col_sample}'],
                    'Frequency': ['N/A']
                })
            
            # Find job codes in dfpurifiers that are not in filtered_dfpurifierjobs
            dfpurifiers['Job Code'] = dfpurifiers[job_code_col].astype(str)
            filtered_dfpurifierjobs['Job Codecopy'] = filtered_dfpurifierjobs['Job Codecopy'].astype(str)
            
            print(f"Sample reference job codes: {dfpurifiers['Job Code'].iloc[:5].tolist()}")
            print(f"Sample current job codes: {filtered_dfpurifierjobs['Job Codecopy'].iloc[:5].tolist()}")
            print(f"Total reference jobs: {len(dfpurifiers)}")
            print(f"Total current purifier jobs: {len(filtered_dfpurifierjobs)}")
            
            missingjobspurifierresult = dfpurifiers[~dfpurifiers['Job Code'].isin(filtered_dfpurifierjobs['Job Codecopy'])]
            print(f"Total missing jobs found: {len(missingjobspurifierresult)}")
            
            # Drop Remarks column if it exists
            if 'Remarks' in missingjobspurifierresult.columns:
                missingjobspurifierresult = missingjobspurifierresult.drop(columns=['Remarks'])
                
            # Reset the index of the combined DataFrame
            missingjobspurifierresult.reset_index(drop=True, inplace=True)
            
            return missingjobspurifierresult
                
        except Exception as e:
            print(f"Error processing purifier reference data: {str(e)}")
            # Create an error row for display
            error_row = pd.DataFrame({
                'Job Code': ['Error'],
                'Title': [f'Unable to process reference data: {str(e)}'],
                'Frequency': ['N/A']
            })
            return error_row