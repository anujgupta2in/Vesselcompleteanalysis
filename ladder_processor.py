import pandas as pd
import numpy as np

class LadderSystemProcessor:
    def __init__(self):
        self.filtered_dfLadderjobs = pd.DataFrame()
        self.result_dfLadder = pd.DataFrame()
        self.pivot_table_resultLadderJobs = pd.DataFrame()
        self.styled_pivot_table_resultLadderJobs = None
        self.missingjobsLadderresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_ladder_data(self, df, dfLadders):
        try:
            # Filter rows where 'Machinery Location' contains 'Ladder'
            self.filtered_dfLadderjobs = df[df['Machinery Location'].str.contains('Ladder', na=False)].copy()

            # Ensure 'Job Codecopy' exists and is string
            if 'Job Codecopy' not in self.filtered_dfLadderjobs.columns:
                self.filtered_dfLadderjobs['Job Codecopy'] = self.filtered_dfLadderjobs['Job Code'].astype(str)
            self.filtered_dfLadderjobs['Job Codecopy'] = self.filtered_dfLadderjobs['Job Codecopy'].astype(object)
            self.filtered_dfLadderjobs['Job Codecopy'] = self.filtered_dfLadderjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            # Ensure both columns used in merge are string
            self.filtered_dfLadderjobs['Job Codecopy'] = self.filtered_dfLadderjobs['Job Codecopy'].astype(str)
            dfLadders['UI Job Code'] = dfLadders['UI Job Code'].astype(str)

            # Merge based on job codes
            self.result_dfLadder = self.filtered_dfLadderjobs.merge(
                dfLadders,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfLadder.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfLadder.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged data for pivot index.")

            # Create pivot table
            pivot_table = self.result_dfLadder.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultLadderJobs = pivot_table
            self.styled_pivot_table_resultLadderJobs = self.pivot_table_resultLadderJobs.style.set_table_styles([
                {'selector': 'th', 'props': [('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('text-align', 'center'), ('min-width', '150px')]},
                {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
            ], overwrite=False)

            # Find missing jobs
            self.missingjobsLadderresult = dfLadders[~dfLadders['UI Job Code'].isin(
                self.filtered_dfLadderjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsLadderresult.columns:
                self.missingjobsLadderresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsLadderresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultLadderJobs = pd.DataFrame({'Error': [f'Ladder data processing failed: {str(e)}']})
            self.missingjobsLadderresult = pd.DataFrame()