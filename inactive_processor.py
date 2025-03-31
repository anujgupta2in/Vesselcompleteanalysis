import pandas as pd
import numpy as np

class InactiveMappingProcessor:
    def __init__(self):
        self.result_dfinactive = pd.DataFrame()
        self.pivot_table_resultinactiveJobs = pd.DataFrame()
        self.pivot_table_resultinactiveJobstotal = pd.DataFrame()
        self.missinginactivejobsresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val in [0, '0', 0.0] else val

    def find_column(self, df, possible_names):
        for name in possible_names:
            if name in df.columns:
                return name
        raise KeyError(f"None of the columns {possible_names} found in DataFrame")

    def process_inactive_data(self, df, dfinactive):
        try:
            dfcopy = df.copy()
            dfcopy['Job Codecopy'] = dfcopy['Job Code'].apply(self.safe_convert_to_string)

            # Identify correct reference column from possible options
            ref_code_col = self.find_column(dfinactive, ['Job Code', 'UI Job Code', 'Code'])
            dfinactive[ref_code_col] = dfinactive[ref_code_col].apply(self.safe_convert_to_string)

            self.result_dfinactive = dfcopy.merge(
                dfinactive,
                left_on='Job Codecopy',
                right_on=ref_code_col,
                suffixes=('_filtered', '_ref')
            )
            self.result_dfinactive.reset_index(drop=True, inplace=True)

            self.result_dfinactive['Title'] = self.result_dfinactive['Title'].apply(lambda x: f"{x:<50}")

            pivot_raw = self.result_dfinactive.pivot_table(
                index='Title',
                columns='Function',
                values='Job Codecopy',
                aggfunc='count'
            ).fillna(0).astype(int)

            self.pivot_table_resultinactiveJobs = pivot_raw.replace(0, '').map(self.format_blank)

            total_raw = self.result_dfinactive.pivot_table(
                index='Title',
                values='Job Codecopy',
                aggfunc='count'
            ).fillna(0).astype(int).sort_values(by='Job Codecopy', ascending=False)

            self.pivot_table_resultinactiveJobstotal = total_raw.replace(0, '').map(self.format_blank)

            self.missinginactivejobsresult = dfinactive[~dfinactive[ref_code_col].isin(dfcopy['Job Codecopy'])].copy()
            self.missinginactivejobsresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultinactiveJobs = pd.DataFrame({'Error': [f'Processing failed: {str(e)}']})
            self.pivot_table_resultinactiveJobstotal = pd.DataFrame()
            self.missinginactivejobsresult = pd.DataFrame()
