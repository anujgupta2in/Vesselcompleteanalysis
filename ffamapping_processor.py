import pandas as pd
import numpy as np

class FFAMappingProcessor:
    def __init__(self):
        self.result_dfffa = pd.DataFrame()
        self.pivot_table_resultffaJobs = pd.DataFrame()
        self.pivot_table_resultffaJobstotal = pd.DataFrame()
        self.missingffajobsresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val in [0, '0', 0.0] else val

    def process_ffa_data(self, df, dfffa):
        try:
            ffa = ['FFE Fixed', 'LSA Fixed', 'LSA Loose', 'FFE Loose']

            dfcopy = df.copy()
            dfcopy['Job Codecopy'] = dfcopy['Job Code'].apply(self.safe_convert_to_string)

            ref_code_col = 'UI Job Code'
            dfffa[ref_code_col] = dfffa[ref_code_col].apply(self.safe_convert_to_string)

            filtered_dfffajobs = dfcopy[dfcopy['Function'].str.contains('|'.join(ffa), na=False)].copy()

            self.result_dfffa = filtered_dfffajobs.merge(
                dfffa,
                left_on='Job Codecopy',
                right_on=ref_code_col,
                suffixes=('_filtered', '_ref')
            )
            self.result_dfffa.reset_index(drop=True, inplace=True)

            # Widen Title column by padding to improve display in Streamlit
            self.result_dfffa['Title'] = self.result_dfffa['Title'].apply(lambda x: f"{x:<50}")

            pivot_raw = self.result_dfffa.pivot_table(
                index='Title',
                columns='Function',
                values='Job Codecopy',
                aggfunc='count'
            ).fillna(0).astype(int)

            self.pivot_table_resultffaJobs = pivot_raw.replace(0, '').map(self.format_blank)

            total_raw = self.result_dfffa.pivot_table(
                index='Title',
                values='Job Codecopy',
                aggfunc='count'
            ).fillna(0).astype(int).sort_values(by='Job Codecopy', ascending=False)

            self.pivot_table_resultffaJobstotal = total_raw.replace(0, '').map(self.format_blank)

            self.missingffajobsresult = dfffa[~dfffa[ref_code_col].isin(filtered_dfffajobs['Job Codecopy'])].copy()
            self.missingffajobsresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultffaJobs = pd.DataFrame({'Error': [f'Processing failed: {str(e)}']})
            self.pivot_table_resultffaJobstotal = pd.DataFrame()
            self.missingffajobsresult = pd.DataFrame()