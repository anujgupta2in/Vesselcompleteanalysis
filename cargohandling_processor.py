import pandas as pd
import numpy as np
import re

class CargoHandlingSystemProcessor:
    def __init__(self):
        self.filter_cargohandling_jobs = pd.DataFrame()
        self.df_cargohanding = pd.DataFrame()
        self.result_df_cargohandling = pd.DataFrame()
        self.missing_jobs_cargohandling = pd.DataFrame()
        self.cargohanding_sheet = ""

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def color_format(self, val):
        return "background-color: #ffd" if val != "" else ""

    def process_reference_data(self, data, ref_sheet):
        try:
            data_copy = data.copy()

            # Step 1: Filter relevant jobs
            pattern = 'Cargo Handling System'
            self.filter_cargohandling_jobs = data_copy[
                data_copy['Function'].str.contains(pattern, na=False, flags=re.IGNORECASE)
            ].copy()

            # Copy job code safely
            self.filter_cargohandling_jobs['Job Codecopy'] = self.filter_cargohandling_jobs['Job Code'].apply(self.safe_convert_to_string)

            # Step 2: Load the reference sheet
            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            self.cargohanding_sheet = 'Cargohanding' if 'Cargohanding' in ref_sheet_names else ref_sheet_names[0]
            self.df_cargohanding = pd.read_excel(ref_sheet, sheet_name=self.cargohanding_sheet)

            # Step 3: Detect usable job code column
            possible_code_cols = ['UI Job Code', 'Job Code', 'JobCode', 'Code']
            job_code_col = next((col for col in possible_code_cols if col in self.df_cargohanding.columns), None)

            if job_code_col is None:
                self.result_df_cargohandling = pd.DataFrame()
                return pd.DataFrame({
                    'Job Code': ['Reference Error'],
                    'Title': ['No recognizable job code column found'],
                    'Frequency': ['N/A']
                })

            # Step 4: Format job codes in reference
            self.df_cargohanding[job_code_col] = (
                self.df_cargohanding[job_code_col]
                .astype(str)
                .str.strip()
                .str.replace(r'\.0$', '', regex=True)
            )

            input_codes = self.filter_cargohandling_jobs['Job Codecopy'] if not self.filter_cargohandling_jobs.empty else []

            # Step 5: Find missing jobs
            self.missing_jobs_cargohandling = self.df_cargohanding[
                ~self.df_cargohanding[job_code_col].isin(input_codes)
            ].reset_index(drop=True)

            # Step 6: Merge matches
            if not self.filter_cargohandling_jobs.empty:
                self.result_df_cargohandling = self.filter_cargohandling_jobs.merge(
                    self.df_cargohanding,
                    left_on='Job Codecopy',
                    right_on=job_code_col,
                    suffixes=('_filtered', '_ref')
                ).reset_index(drop=True)
            else:
                self.result_df_cargohandling = pd.DataFrame()

            return self.result_df_cargohandling

        except Exception as e:
            self.result_df_cargohandling = pd.DataFrame()
            return pd.DataFrame({
                'Job Code': ['Error'],
                'Title': [str(e)],
                'Frequency': ['N/A']
            })

    def create_task_count_table(self):
        try:
            if self.result_df_cargohandling.empty:
                return pd.DataFrame({'Error': ['No matching Cargo Handling System jobs to analyze.']})

            if 'Machinery Locationcopy' not in self.result_df_cargohandling.columns:
                self.result_df_cargohandling['Machinery Locationcopy'] = self.result_df_cargohandling.get('Machinery Location', '')

            pivot = self.result_df_cargohandling.pivot_table(
                index='Title',
                columns='Machinery Locationcopy',
                values='Job Codecopy',
                aggfunc='count'
            ).replace(np.nan, '').replace('', -1).astype(int).map(self.format_blank)

            return pivot.style.map(self.color_format)

        except Exception as e:
            return pd.DataFrame({'Error': [f'Task count table creation failed: {str(e)}']})

    def get_results_dict(self):
        return {
            "matching_jobs": self.result_df_cargohandling,
            "missing_jobs": self.missing_jobs_cargohandling
        }
