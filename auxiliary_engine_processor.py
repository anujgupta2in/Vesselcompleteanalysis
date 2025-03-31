import pandas as pd
import numpy as np

class AuxiliaryEngineProcessor:
    def __init__(self):
        """Initialize AuxiliaryEngineProcessor with component list."""
        self.components_to_check = [
            'Turbocharger - AE',
            'Air Cooler - AE',
            'Alternator - AE',
            'Connecting Rod - AE',
            'Cylinder Head - AE',
            'Fuel Injection Pump - AE',
            'Fuel Valve - AE',
            'Governor - AE',
            'Main Bearing - AE',
            'Piston - AE',
            'Attached LO Pump - AE',
            'FO Filter - AE',
            'LO Cooler - AE',
            'LO Filter - AE',
            'LT Attached CW Pump - AE',
        ]
        self.job_codes = [
            ([6619, 2157], "Piston Overhaul", r'Piston - AE#(\d+)'),
            ([2222, 2223, 2224], "Cylinder Head Overhaul", r'Cylinder Head - AE#(\d+)'),
            ([2196, 2191], "Fuel Valve Overhaul", r'Fuel Valve - AE#(\d+)'),
            (2202, "Fuel Pump Overhaul", r'Fuel Injection Pump - AE#(\d+)'),
            ([3878, 2127, 2126], "TurboCharger Overhaul", r'Auxiliary Engine#(\d+)')
        ]

    def extract_running_hours(self, data):
        """Extract running hours for auxiliary engines."""
        aux_running_hours = {'AE1': "Not Available", 'AE2': "Not Available", 'AE3': "Not Available"}
        if 'Machinery Running Hours' in data.columns:
            for ae_num in range(1, 4):
                ae_data = data[data['Machinery Location'].str.contains(f'Auxiliary Engine#{ae_num}', case=False, na=False)]
                if not ae_data.empty:
                    running_hours = ae_data['Machinery Running Hours'].dropna()
                    if not running_hours.empty:
                        aux_running_hours[f'AE{ae_num}'] = str(int(float(running_hours.iloc[0])))
        return aux_running_hours

    def format_unit_data(self, unit_data):
        """Format unit data for display."""
        try:
            last_done_date = unit_data['Last Done Date'].dropna().astype(str).iloc[0] if not unit_data['Last Done Date'].dropna().empty else "No Date"
            last_done_running_hours = unit_data['Last Done Running Hours'].dropna().astype(str).iloc[0] if not unit_data['Last Done Running Hours'].dropna().empty else "No RH"
            remaining_hours = unit_data['Remaining Running Hours'].dropna().astype(str).iloc[0] if not unit_data['Remaining Running Hours'].dropna().empty else "No RH"
            return f"Date: {last_done_date}\nRH: {last_done_running_hours}\nRemaining Hours: {remaining_hours}"
        except Exception:
            return "Data formatting error"

    def process_job_code(self, data, job_codes, job_description, unit_pattern):
        """Process job codes for auxiliary engines."""
        if not isinstance(job_codes, list):
            job_codes = [job_codes]
        try:
            job_data = data[data['Job Code'].isin(job_codes)][
                ['Frequency', 'Last Done Date', 'Last Done Running Hours', 
                 'Remaining Running Hours', 'Machinery Location', 'Sub Component Location']
            ].copy()
            job_data['Unit'] = job_data['Machinery Location'].str.extract(r'Auxiliary Engine#(\d+)')

            structured_data = pd.DataFrame({
                'Job Title': [job_description],
                'Frequency': [job_data['Frequency'].iloc[0] if not job_data.empty else "No Frequency"]
            })

            for unit in range(1, 4):  # AE1, AE2, AE3
                unit_str = str(unit)
                unit_data = job_data[job_data['Unit'] == unit_str]
                if not unit_data.empty:
                    structured_data[f'AE{unit}'] = [self.format_unit_data(unit_data)]
                else:
                    structured_data[f'AE{unit}'] = ["No Data Available"]

            return structured_data
        except Exception as e:
            print(f"Error processing AE job code {job_codes}: {str(e)}")
            return pd.DataFrame({'Job Title': [job_description], 'Frequency': ["Error processing data"]})

    def get_maintenance_data(self, data):
        """Get auxiliary engine maintenance data."""
        maintenance_data = []
        for job_info in self.job_codes:
            if len(job_info) == 3:
                job_codes, job_description, unit_pattern = job_info
                structured_data = self.process_job_code(data, job_codes, job_description, unit_pattern)
                maintenance_data.append(structured_data)

        if maintenance_data:
            return pd.concat(maintenance_data, ignore_index=True)
        return pd.DataFrame()

    def analyze_components(self, data):
        """Analyze component presence for auxiliary engines."""
        appended_set = set(data['Machinery Location'].fillna('')) | set(data['Sub Component Location'].fillna(''))
        
        component_list = []
        status_list = []
        missing_count = 0
        
        for component in self.components_to_check:
            found = any(component in item for item in appended_set)
            component_list.append(component)
            status_list.append("Present" if found else "Missing")
            if not found:
                missing_count += 1
        
        component_status = pd.DataFrame({
            'Component': component_list,
            'Status': status_list
        })
        
        return component_status, missing_count

    def create_task_count_table(self, data):
        """Create task count analysis table for auxiliary engines."""
        try:
            # Filter auxiliary engine data using the regex pattern
            auxiliary_engine_df = data[data['Machinery Location'].str.contains('Auxiliary Engine(?:No|#)[1-4]', na=False, regex=True)]

            # Create basic task count pivot table
            pivot_table = auxiliary_engine_df.pivot_table(
                index='Machinery Location',
                values='Title',
                aggfunc='count'
            ).reset_index()

            # Rename columns for clarity
            pivot_table.columns = ['Machinery Location', 'Task Count']

            return pivot_table

        except Exception as e:
            print(f"Error creating task count table: {e}")
            return None

    def create_component_distribution(self, data):
        """Create component distribution analysis table."""
        if 'Sub Component Location' not in data.columns:
            return None

        pattern = '|'.join(self.components_to_check)
        filtered_df = data[data['Sub Component Location'].str.contains(pattern, na=False, case=False)].copy()
        pivot_table = filtered_df.pivot_table(
            index=['Sub Component Location', 'Title'],
            columns='Machinery Location',
            values='Job Code',
            aggfunc='count',
            fill_value=0
        ).reset_index()
        pivot_table.columns.name = None
        return pivot_table

    def process_reference_data(self, data, ref_sheet):
        """Process reference data for auxiliary engines."""
        try:
            filtered_df = data[data['Machinery Location'].str.contains("Auxiliary Engine", na=False, case=False)].copy()
            filtered_df['Job Codecopy'] = filtered_df['Job Code'].astype(str)

            ref_df = pd.read_excel(ref_sheet, sheet_name='AE Jobs')
            ref_df['UI Job Code'] = ref_df['UI Job Code'].astype(str)

            result_df = filtered_df.merge(
                ref_df,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )

            pivot_table = result_df.pivot_table(
                index='Title',
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count',
                fill_value=0
            ).reset_index()
            pivot_table.columns.name = None

            missing_jobs = ref_df[~ref_df['UI Job Code'].isin(filtered_df['Job Codecopy'])]
            if 'Remarks' in missing_jobs.columns:
                missing_jobs = missing_jobs.drop(columns=['Remarks'])
            missing_jobs.reset_index(drop=True, inplace=True)

            return pivot_table, missing_jobs
        except Exception as e:
            print(f"Error processing AE reference data: {e}")
            return None, None