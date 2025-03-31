import pandas as pd
import numpy as np
import re

class CargoPumpingProcessor:
    def __init__(self):
        self.filtercargo_pumping_jobs = pd.DataFrame()
        self.dfcargopumping = pd.DataFrame()
        self.missingjobscargopumpingresult = pd.DataFrame()
        self.result_dfcargopumping = pd.DataFrame()
        self.cargopumping_sheet = ""

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def color_format(self, val):
        return "background-color: #ffd" if val != "" else ""

    def format_unit_data(self, unit_data):
        try:
            formatted_data = []
            for unit, tasks in unit_data.items():
                task_text = "<br>".join(tasks) if isinstance(tasks, list) else str(tasks)
                formatted_data.append({
                    "Cargo Pumping System": unit,
                    "Maintenance Tasks": task_text
                })
            return pd.DataFrame(formatted_data)
        except Exception as e:
            print(f"Error formatting Cargo Pumping System data: {str(e)}")
            return pd.DataFrame(columns=['Cargo Pumping System', 'Maintenance Tasks'])

    def process_reference_data(self, data, ref_sheet):
        try:
            data_copy = data.copy()

            # Flexible filtering: include broader cargo-related systems
            pattern = 'Cargo Pumping|Cargo Oil|Cargo'
            self.filtercargo_pumping_jobs = data_copy[data_copy['Function'].str.contains(pattern, na=False, flags=re.IGNORECASE)].copy()

            if self.filtercargo_pumping_jobs.empty:
                return pd.DataFrame({'Job Code': ['No Cargo Pumping jobs found'], 'Title': ['Check Function values'], 'Frequency': ['N/A']})

            # Ensure 'Job Codecopy' is string
            self.filtercargo_pumping_jobs['Job Codecopy'] = self.filtercargo_pumping_jobs['Job Code'].apply(self.safe_convert_to_string)

            # Read reference sheet
            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            self.cargopumping_sheet = 'Cargo Pumping' if 'Cargo Pumping' in ref_sheet_names else ref_sheet_names[0]
            self.dfcargopumping = pd.read_excel(ref_sheet, sheet_name=self.cargopumping_sheet)

            if 'UI Job Code' not in self.dfcargopumping.columns:
                return pd.DataFrame({'Job Code': ['Reference Error'], 'Title': ['Missing UI Job Code'], 'Frequency': ['N/A']})

            self.dfcargopumping['UI Job Code'] = self.dfcargopumping['UI Job Code'].astype(str).str.strip()

            # Merge data
            self.result_dfcargopumping = self.filtercargo_pumping_jobs.merge(
                self.dfcargopumping,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            ).reset_index(drop=True)

            # Find missing jobs
            self.missingjobscargopumpingresult = self.dfcargopumping[~self.dfcargopumping['UI Job Code'].isin(
                self.filtercargo_pumping_jobs['Job Codecopy'])].reset_index(drop=True)

            return self.result_dfcargopumping[['UI Job Code', 'J3 Job Title', 'Remarks', 'Applicability']] if 'J3 Job Title' in self.result_dfcargopumping.columns else self.result_dfcargopumping

        except Exception as e:
            print(f"Error processing Cargo Pumping data: {str(e)}")
            return pd.DataFrame({'Job Code': ['Error'], 'Title': [str(e)], 'Frequency': ['N/A']})

    def create_task_count_table(self):
        try:
            if self.result_dfcargopumping.empty:
                return pd.DataFrame({'Error': ['No matching Cargo Pumping jobs to analyze.']})

            if 'Machinery Locationcopy' not in self.result_dfcargopumping.columns:
                self.result_dfcargopumping['Machinery Locationcopy'] = self.result_dfcargopumping.get('Machinery Location', '')

            pivot = self.result_dfcargopumping.pivot_table(
                index='Title',
                columns='Machinery Locationcopy',
                values='Job Codecopy',
                aggfunc='count'
            ).replace(np.nan, '').replace('', -1).astype(int).map(self.format_blank)

            return pivot.style.map(self.color_format)

        except Exception as e:
            return pd.DataFrame({'Error': [f'Task count table creation failed: {str(e)}']})
