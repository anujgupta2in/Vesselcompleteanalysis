import pandas as pd
import numpy as np

class BTSystemProcessor:
    def __init__(self):
        self.filtered_dfBTjobs = pd.DataFrame()
        self.result_dfBT = pd.DataFrame()
        self.pivot_table_resultBTJobs = pd.DataFrame()
        self.styled_pivot_table_resultBTJobs = None
        self.missingjobsBTresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_bt_data(self, df, dfBT):
        try:
            # Filter data
            BT = ['Bow Thruster']
            self.filtered_dfBTjobs = df[df['Machinery Location'].str.contains('|'.join(BT), na=False)].copy()

            # Ensure 'Job Codecopy' column
            if 'Job Codecopy' not in self.filtered_dfBTjobs.columns:
                self.filtered_dfBTjobs['Job Codecopy'] = self.filtered_dfBTjobs['Job Code'].astype(str)
            self.filtered_dfBTjobs['Job Codecopy'] = self.filtered_dfBTjobs['Job Codecopy'].astype(object)
            self.filtered_dfBTjobs['Job Codecopy'] = self.filtered_dfBTjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dfBT['UI Job Code'] = dfBT['UI Job Code'].astype(str)

            # Merge
            self.result_dfBT = self.filtered_dfBTjobs.merge(
                dfBT,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfBT.reset_index(drop=True, inplace=True)

            # Pivot table
            title_col = next((col for col in ['Title', 'J3 Job Title', 'Task Description', 'Job Title'] if col in self.result_dfBT.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged BT data for pivot index.")

            self.pivot_table_resultBTJobs = self.result_dfBT.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            self.pivot_table_resultBTJobs.replace(np.nan, '', inplace=True)
            self.pivot_table_resultBTJobs.replace('', -1, inplace=True)
            self.pivot_table_resultBTJobs = self.pivot_table_resultBTJobs.astype(int)
            self.pivot_table_resultBTJobs = self.pivot_table_resultBTJobs.map(self.format_blank)

            self.styled_pivot_table_resultBTJobs = self.pivot_table_resultBTJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Missing jobs
            self.missingjobsBTresult = dfBT[~dfBT['UI Job Code'].isin(
                self.filtered_dfBTjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsBTresult.columns:
                self.missingjobsBTresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsBTresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultBTJobs = pd.DataFrame({'Error': [f'BT data processing failed: {str(e)}']})
            self.missingjobsBTresult = pd.DataFrame()
