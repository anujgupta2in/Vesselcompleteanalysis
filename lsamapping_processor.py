import pandas as pd
import numpy as np

class LSAMappingProcessor:
    def __init__(self):
        self.result_dflsa = pd.DataFrame()
        self.pivot_table_resultlsaJobs = pd.DataFrame()
        self.pivot_table_resultlsaJobstotal = pd.DataFrame()
        self.missinglsajobsresult = pd.DataFrame()
        self.pivot_table_resultlsaJobs_styled = None  # Optional: for UI display

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val in [0, '0', 0.0] else val

    def highlight_one(self, val):
        if val == 1:
            return 'background-color: #ffe599'
        return ''

    def process_lsa_data(self, df, dflsa):
        try:
            dfcopy = df.copy()
            dfcopy['Job Codecopy'] = dfcopy['Job Code'].apply(self.safe_convert_to_string)
            dflsa['UI Job Code'] = dflsa['UI Job Code'].apply(self.safe_convert_to_string)

            self.result_dflsa = dfcopy.merge(
                dflsa,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dflsa.reset_index(drop=True, inplace=True)

            # Widen Title column for better display (optional)
            self.result_dflsa['Title'] = self.result_dflsa['Title'].apply(lambda x: f"{x:<50}")

            pivot_raw = self.result_dflsa.pivot_table(
                index='Title',
                columns='Function',
                values='Job Codecopy',
                aggfunc='count'
            ).fillna(0).astype(int)

            self.pivot_table_resultlsaJobs = pivot_raw.replace(0, '').map(self.format_blank)

            # Optional: keep styled version for UI display
            self.pivot_table_resultlsaJobs_styled = pivot_raw.style.applymap(self.highlight_one)

            total_raw = self.result_dflsa.pivot_table(
                index='Title',
                values='Job Codecopy',
                aggfunc='count'
            ).fillna(0).astype(int).sort_values(by='Job Codecopy', ascending=False)

            self.pivot_table_resultlsaJobstotal = total_raw.replace(0, '').map(self.format_blank)

            self.missinglsajobsresult = dflsa[~dflsa['UI Job Code'].isin(dfcopy['Job Codecopy'])].copy()
            self.missinglsajobsresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultlsaJobs = pd.DataFrame({'Error': [f'Processing failed: {str(e)}']})
            self.pivot_table_resultlsaJobstotal = pd.DataFrame()
            self.missinglsajobsresult = pd.DataFrame()
