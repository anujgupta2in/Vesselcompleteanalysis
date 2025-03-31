import pandas as pd
import numpy as np

class CriticalJobsProcessor:
    def __init__(self):
        self.result_dfcritical = pd.DataFrame()
        self.pivot_table_resultcriticalJobs = pd.DataFrame()
        self.pivot_table_resultcriticalJobstotal = pd.DataFrame()
        self.missingcriticaljobsresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val in [0, '0', 0.0] else val

    def process_critical_data(self, df, dfcritical):
        try:
            dfcopy = df.copy()
            dfcopy['Job Codecopy'] = dfcopy['Job Code'].apply(self.safe_convert_to_string)

            # Dynamically find the correct column name for job codes in dfcritical
            possible_job_code_cols = ['Job Code', 'UI Job Code', 'Code']
            ref_code_col = next((col for col in possible_job_code_cols if col in dfcritical.columns), None)
            if not ref_code_col:
                raise KeyError("No valid Job Code column found in critical reference sheet.")

            dfcritical[ref_code_col] = dfcritical[ref_code_col].apply(self.safe_convert_to_string)

            self.result_dfcritical = dfcopy.merge(
                dfcritical,
                left_on='Job Codecopy',
                right_on=ref_code_col,
                suffixes=('_filtered', '_ref')
            )
            self.result_dfcritical.reset_index(drop=True, inplace=True)

            self.result_dfcritical['Title'] = self.result_dfcritical['Title'].apply(lambda x: f"{x:<50}")

            pivot_raw = self.result_dfcritical.pivot_table(
                index='Title',
                columns='Function',
                values='Job Codecopy',
                aggfunc='count'
            ).fillna(0).astype(int)

            self.pivot_table_resultcriticalJobs = pivot_raw.replace(0, '').map(self.format_blank)

            total_raw = self.result_dfcritical.pivot_table(
                index='Title',
                values='Job Codecopy',
                aggfunc='count'
            ).fillna(0).astype(int).sort_values(by='Job Codecopy', ascending=False)

            self.pivot_table_resultcriticalJobstotal = total_raw.replace(0, '').map(self.format_blank)

            self.missingcriticaljobsresult = dfcritical[~dfcritical[ref_code_col].isin(dfcopy['Job Codecopy'])].copy()
            self.missingcriticaljobsresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultcriticalJobs = pd.DataFrame({'Error': [f'Processing failed: {str(e)}']})
            self.pivot_table_resultcriticalJobstotal = pd.DataFrame()
            self.missingcriticaljobsresult = pd.DataFrame()
