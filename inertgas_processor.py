import pandas as pd
import numpy as np
import re

class InertGasSystemProcessor:
    def __init__(self):
        self.filter_igsystem_jobs = pd.DataFrame()
        self.df_igsystem = pd.DataFrame()
        self.result_df_igsystem = pd.DataFrame()
        self.missing_jobs_igsystem = pd.DataFrame()
        self.igsystem_sheet = ""

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def color_format(self, val):
        return "background-color: #ffd" if val != "" else ""

    def get_results_dict(self):
        return {
            "matching_jobs": self.result_df_igsystem,
            "missing_jobs": self.missing_jobs_igsystem
        }

    def process_reference_data(self, data, ref_sheet):
        try:
            data_copy = data.copy()
            pattern = 'Inert Gas|IG '
            self.filter_igsystem_jobs = data_copy[data_copy['Function'].str.contains(pattern, na=False, flags=re.IGNORECASE)].copy()
            self.filter_igsystem_jobs['Job Codecopy'] = self.filter_igsystem_jobs['Job Code'].apply(self.safe_convert_to_string)

            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            self.igsystem_sheet = 'IGSystem' if 'IGSystem' in ref_sheet_names else ref_sheet_names[0]
            self.df_igsystem = pd.read_excel(ref_sheet, sheet_name=self.igsystem_sheet)

            possible_code_cols = ['UI Job Code', 'Job Code', 'JobCode', 'Code']
            job_code_col = next((col for col in possible_code_cols if col in self.df_igsystem.columns), None)

            if job_code_col is None:
                self.result_df_igsystem = pd.DataFrame()
                return pd.DataFrame({'Job Code': ['Reference Error'], 'Title': ['No recognizable job code column found'], 'Frequency': ['N/A']})

            self.df_igsystem[job_code_col] = (
                self.df_igsystem[job_code_col]
                .astype(str)
                .str.strip()
                .str.replace(r'\.0$', '', regex=True)
            )

            input_codes = self.filter_igsystem_jobs['Job Codecopy'] if not self.filter_igsystem_jobs.empty else []
            self.missing_jobs_igsystem = self.df_igsystem[~self.df_igsystem[job_code_col].isin(input_codes)].reset_index(drop=True)

            if not self.filter_igsystem_jobs.empty:
                self.result_df_igsystem = self.filter_igsystem_jobs.merge(
                    self.df_igsystem,
                    left_on='Job Codecopy',
                    right_on=job_code_col,
                    suffixes=('_filtered', '_ref')
                ).reset_index(drop=True)
            else:
                self.result_df_igsystem = pd.DataFrame()

            return self.result_df_igsystem

        except Exception as e:
            self.result_df_igsystem = pd.DataFrame()
            return pd.DataFrame({'Job Code': ['Error'], 'Title': [str(e)], 'Frequency': ['N/A']})

    def create_task_count_table(self):
        try:
            if self.result_df_igsystem.empty:
                return pd.DataFrame({'Error': ['No matching Inert Gas System jobs to analyze.']})

            if 'Machinery Locationcopy' not in self.result_df_igsystem.columns:
                self.result_df_igsystem['Machinery Locationcopy'] = self.result_df_igsystem.get('Machinery Location', '')

            pivot = self.result_df_igsystem.pivot_table(
                index='Title',
                columns='Machinery Locationcopy',
                values='Job Codecopy',
                aggfunc='count'
            ).replace(np.nan, '').replace('', -1).astype(int).map(self.format_blank)

            return pivot.style.map(self.color_format)

        except Exception as e:
            return pd.DataFrame({'Error': [f'Task count table creation failed: {str(e)}']})
