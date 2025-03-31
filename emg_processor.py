import pandas as pd
import numpy as np

class EmergencyGenSystemProcessor:
    def __init__(self):
        self.filtered_dfEmgjobs = pd.DataFrame()
        self.result_dfEmg = pd.DataFrame()
        self.pivot_table_resultEmgJobs = pd.DataFrame()
        self.styled_pivot_table_resultEmgJobs = None
        self.missingjobsEmgresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_emg_data(self, df, dfEmg):
        try:
            Emg = ['Emergency Gen']

            self.filtered_dfEmgjobs = df[df['Machinery Location'].str.contains('|'.join(Emg), na=False)].copy()

            if 'Job Codecopy' not in self.filtered_dfEmgjobs.columns:
                self.filtered_dfEmgjobs['Job Codecopy'] = self.filtered_dfEmgjobs['Job Code'].astype(str)
            self.filtered_dfEmgjobs['Job Codecopy'] = self.filtered_dfEmgjobs['Job Codecopy'].astype(object)
            self.filtered_dfEmgjobs['Job Codecopy'] = self.filtered_dfEmgjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dfEmg['UI Job Code'] = dfEmg['UI Job Code'].astype(str)

            self.result_dfEmg = self.filtered_dfEmgjobs.merge(
                dfEmg,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfEmg.reset_index(drop=True, inplace=True)

            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfEmg.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged Emergency Gen data for pivot index.")

            pivot_table = self.result_dfEmg.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultEmgJobs = pivot_table
            self.styled_pivot_table_resultEmgJobs = self.pivot_table_resultEmgJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.missingjobsEmgresult = dfEmg[~dfEmg['UI Job Code'].isin(
                self.filtered_dfEmgjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsEmgresult.columns:
                self.missingjobsEmgresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsEmgresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultEmgJobs = pd.DataFrame({'Error': [f'Emergency Generator data processing failed: {str(e)}']})
            self.missingjobsEmgresult = pd.DataFrame()
