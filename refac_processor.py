import pandas as pd
import numpy as np

class RefacSystemProcessor:
    def __init__(self):
        self.filtered_dfrefacjobs = pd.DataFrame()
        self.result_dfrefac = pd.DataFrame()
        self.pivot_table_resultrefacJobs = pd.DataFrame()
        self.styled_pivot_table_resultrefacJobs = None
        self.missingjobsrefacresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_refac_data(self, df, dfrefac):
        try:
            refac = ['AC Plant', 'Air Handling Unit ', 'Packaged AC', 'Refrigeration Plant', 'Refrigeration System']

            self.filtered_dfrefacjobs = df[df['Machinery Location'].str.contains('|'.join(refac), na=False)].copy()

            if 'Job Codecopy' not in self.filtered_dfrefacjobs.columns:
                self.filtered_dfrefacjobs['Job Codecopy'] = self.filtered_dfrefacjobs['Job Code'].astype(str)
            self.filtered_dfrefacjobs['Job Codecopy'] = self.filtered_dfrefacjobs['Job Codecopy'].astype(object)
            self.filtered_dfrefacjobs['Job Codecopy'] = self.filtered_dfrefacjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dfrefac['UI Job Code'] = dfrefac['UI Job Code'].astype(str)

            self.result_dfrefac = self.filtered_dfrefacjobs.merge(
                dfrefac,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfrefac.reset_index(drop=True, inplace=True)

            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfrefac.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged refac data for pivot index.")

            pivot_table = self.result_dfrefac.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultrefacJobs = pivot_table
            self.styled_pivot_table_resultrefacJobs = self.pivot_table_resultrefacJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.missingjobsrefacresult = dfrefac[~dfrefac['UI Job Code'].isin(
                self.filtered_dfrefacjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsrefacresult.columns:
                self.missingjobsrefacresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsrefacresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultrefacJobs = pd.DataFrame({'Error': [f'Refac data processing failed: {str(e)}']})
            self.missingjobsrefacresult = pd.DataFrame()
