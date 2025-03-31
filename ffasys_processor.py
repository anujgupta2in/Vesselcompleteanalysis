import pandas as pd
import numpy as np

class FFASystemProcessor:
    def __init__(self):
        self.filter_ffasys_jobs = pd.DataFrame()
        self.df_ffasys = pd.DataFrame()
        self.result_df_ffasys = pd.DataFrame()
        self.missing_jobs_ffasys = pd.DataFrame()
        self.ffasys_sheet = ""

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def process_reference_data(self, data, ref_sheet):
        try:
            data_copy = data.copy()
            ffasys = ['Fire Fighting System']
            pattern = '|'.join(ffasys)
            self.filter_ffasys_jobs = data_copy[data_copy['Machinery Location'].str.contains(pattern, na=False)].copy()
            self.filter_ffasys_jobs['Job Codecopy'] = self.filter_ffasys_jobs['Job Code'].apply(self.safe_convert_to_string)

            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            self.ffasys_sheet = 'FFASYS' if 'FFASYS' in ref_sheet_names else ref_sheet_names[0]
            self.df_ffasys = pd.read_excel(ref_sheet, sheet_name=self.ffasys_sheet)

            job_code_col = 'UI Job Code'
            self.df_ffasys[job_code_col] = self.df_ffasys[job_code_col].astype(str).str.strip()

            input_codes = self.filter_ffasys_jobs['Job Codecopy'] if not self.filter_ffasys_jobs.empty else []
            self.missing_jobs_ffasys = self.df_ffasys[~self.df_ffasys[job_code_col].isin(input_codes)].drop(columns=['Remarks'], errors='ignore').reset_index(drop=True)

            if not self.filter_ffasys_jobs.empty:
                self.result_df_ffasys = self.filter_ffasys_jobs.merge(
                    self.df_ffasys,
                    left_on='Job Codecopy',
                    right_on=job_code_col,
                    suffixes=('_filtered', '_ref')
                ).reset_index(drop=True)
                return self.result_df_ffasys

            return pd.DataFrame({'Job Code': ['No Fire Fighting System jobs found'], 'Title': ['Check Machinery Location values'], 'Frequency': ['N/A']})

        except Exception as e:
            return pd.DataFrame({'Job Code': ['Error'], 'Title': [str(e)], 'Frequency': ['N/A']})

    def create_task_count_table(self):
        try:
            if self.result_df_ffasys.empty:
                return pd.DataFrame({'Error': ['No matching Fire Fighting System jobs to analyze.']})

            if 'Machinery Location' not in self.result_df_ffasys.columns:
                self.result_df_ffasys['Machinery Location'] = ''

            pivot = self.result_df_ffasys.pivot_table(
                index='Title',
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            ).replace(np.nan, '').replace('', -1).astype(int)

            pivot = pivot.map(lambda val: "" if val == -1 else val)
            return pivot.style.map(lambda val: "background-color: #ffd" if val != "" else "")

        except Exception as e:
            return pd.DataFrame({'Error': [f'Task count table creation failed: {str(e)}']})
