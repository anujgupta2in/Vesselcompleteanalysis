import pandas as pd
import numpy as np

class LSAFFAProcessor:
    def __init__(self):
        self.filter_lsaffa_jobs = pd.DataFrame()
        self.df_lsaffa = pd.DataFrame()
        self.result_df_lsaffa = pd.DataFrame()
        self.missing_jobs_lsaffa = pd.DataFrame()
        self.lsaffa_sheet = ""

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def process_reference_data(self, data, ref_sheet):
        try:
            data_copy = data.copy()
            lsaffa = ['FFE Fixed', 'LSA Fixed', 'LSA Loose', 'FFE Loose']
            pattern = '|'.join(lsaffa)
            self.filter_lsaffa_jobs = data_copy[data_copy['Function'].str.contains(pattern, na=False)].copy()
            self.filter_lsaffa_jobs['Job Codecopy'] = self.filter_lsaffa_jobs['Job Code'].apply(self.safe_convert_to_string)

            ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names
            self.lsaffa_sheet = 'LSAFFA' if 'LSAFFA' in ref_sheet_names else ref_sheet_names[0]
            self.df_lsaffa = pd.read_excel(ref_sheet, sheet_name=self.lsaffa_sheet)

            job_code_col = 'UI Job Code'
            self.df_lsaffa[job_code_col] = self.df_lsaffa[job_code_col].astype(str).str.strip()

            input_codes = self.filter_lsaffa_jobs['Job Codecopy'] if not self.filter_lsaffa_jobs.empty else []
            self.missing_jobs_lsaffa = self.df_lsaffa[~self.df_lsaffa[job_code_col].isin(input_codes)].drop(columns=['Remarks'], errors='ignore').reset_index(drop=True)

            if not self.filter_lsaffa_jobs.empty:
                self.result_df_lsaffa = self.filter_lsaffa_jobs.merge(
                    self.df_lsaffa,
                    left_on='Job Codecopy',
                    right_on=job_code_col,
                    suffixes=('_filtered', '_ref')
                ).reset_index(drop=True)
                return self.result_df_lsaffa

            return pd.DataFrame({'Job Code': ['No LSA/FFA jobs found'], 'Title': ['Check Function values'], 'Frequency': ['N/A']})

        except Exception as e:
            return pd.DataFrame({'Job Code': ['Error'], 'Title': [str(e)], 'Frequency': ['N/A']})

    def create_task_count_table(self):
        try:
            if self.result_df_lsaffa.empty:
                return pd.DataFrame({'Error': ['No matching LSA/FFA jobs to analyze.']})

            if 'Function' not in self.result_df_lsaffa.columns:
                self.result_df_lsaffa['Function'] = ''

            pivot = self.result_df_lsaffa.pivot_table(
                index='Title',
                columns='Function',
                values='Job Codecopy',
                aggfunc='count'
            ).replace(np.nan, '').replace('', -1).astype(int)

            pivot = pivot.map(lambda val: "" if val == -1 else val)
            return pivot.style.map(lambda val: "background-color: #ffd" if val != "" else "")

        except Exception as e:
            return pd.DataFrame({'Error': [f'Task count table creation failed: {str(e)}']})
