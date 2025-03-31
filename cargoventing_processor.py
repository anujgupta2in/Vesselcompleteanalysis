import pandas as pd
import numpy as np
import re

class CargoVentingSystemProcessor:
    def __init__(self):
        self.filter_cargovent_jobs = pd.DataFrame()
        self.df_cargovent = pd.DataFrame()
        self.result_df_cargovent = pd.DataFrame()
        self.missing_jobs_cargovent = pd.DataFrame()
        self.cargovent_sheet = ""

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

            # Step 1: Filter jobs
            pattern = 'Cargo Ventilation System'
            self.filter_cargovent_jobs = data_copy[
                data_copy['Function'].str.contains(pattern, na=False, flags=re.IGNORECASE)
            ].copy()
            self.filter_cargovent_jobs['Job Codecopy'] = self.filter_cargovent_jobs['Job Code'].apply(self.safe_convert_to_string)

            # Step 2: Load reference
            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            self.cargovent_sheet = 'Cargovent' if 'Cargovent' in ref_sheet_names else ref_sheet_names[0]
            self.df_cargovent = pd.read_excel(ref_sheet, sheet_name=self.cargovent_sheet)

            # Step 3: Identify job code column
            possible_code_cols = ['UI Job Code', 'Job Code', 'JobCode', 'Code']
            job_code_col = next((col for col in possible_code_cols if col in self.df_cargovent.columns), None)

            if job_code_col is None:
                self.result_df_cargovent = pd.DataFrame()
                return pd.DataFrame({
                    'Job Code': ['Reference Error'],
                    'Title': ['No recognizable job code column found'],
                    'Frequency': ['N/A']
                })

            # Step 4: Format job codes
            self.df_cargovent[job_code_col] = (
                self.df_cargovent[job_code_col]
                .astype(str)
                .str.strip()
                .str.replace(r'\.0$', '', regex=True)
            )

            input_codes = self.filter_cargovent_jobs['Job Codecopy'] if not self.filter_cargovent_jobs.empty else []

            # Step 5: Find missing
            self.missing_jobs_cargovent = self.df_cargovent[
                ~self.df_cargovent[job_code_col].isin(input_codes)
            ].reset_index(drop=True)

            # Step 6: Merge matched
            if not self.filter_cargovent_jobs.empty:
                self.result_df_cargovent = self.filter_cargovent_jobs.merge(
                    self.df_cargovent,
                    left_on='Job Codecopy',
                    right_on=job_code_col,
                    suffixes=('_filtered', '_ref')
                ).reset_index(drop=True)
            else:
                self.result_df_cargovent = pd.DataFrame()

            return self.result_df_cargovent

        except Exception as e:
            self.result_df_cargovent = pd.DataFrame()
            return pd.DataFrame({
                'Job Code': ['Error'],
                'Title': [str(e)],
                'Frequency': ['N/A']
            })

    def create_task_count_table(self):
        try:
            if self.result_df_cargovent.empty:
                return pd.DataFrame({'Error': ['No matching Cargo Ventilation System jobs to analyze.']})

            if 'Machinery Locationcopy' not in self.result_df_cargovent.columns:
                self.result_df_cargovent['Machinery Locationcopy'] = self.result_df_cargovent.get('Machinery Location', '')

            pivot = self.result_df_cargovent.pivot_table(
                index='Title',
                columns='Machinery Locationcopy',
                values='Job Codecopy',
                aggfunc='count'
            ).replace(np.nan, '').replace('', -1).astype(int).map(self.format_blank)

            return pivot.style.map(self.color_format)

        except Exception as e:
            return pd.DataFrame({'Error': [f'Task count table creation failed: {str(e)}']})
