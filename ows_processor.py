import pandas as pd
import numpy as np

class OWSSystemProcessor:
    def __init__(self):
        self.filtered_dfOWS = pd.DataFrame()
        self.result_dfOWS = pd.DataFrame()
        self.pivot_table_resultOWSJobs = pd.DataFrame()
        self.styled_pivot_table_resultOWSJobs = None
        self.missingjobsOWSresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_ows_data(self, df, dfOWS):
        try:
            # Filter rows where 'Machinery Location' contains 'Oily Water'
            self.filtered_dfOWS = df[df['Machinery Location'].str.contains('Oily Water', na=False)].copy()

            # Ensure 'Job Codecopy' is of object type and safely converted to string
            if 'Job Codecopy' not in self.filtered_dfOWS.columns:
                self.filtered_dfOWS['Job Codecopy'] = self.filtered_dfOWS['Job Code'].astype(str)
            self.filtered_dfOWS['Job Codecopy'] = self.filtered_dfOWS['Job Codecopy'].astype(object)
            self.filtered_dfOWS['Job Codecopy'] = self.filtered_dfOWS['Job Codecopy'].apply(self.safe_convert_to_string)

            # Ensure reference codes are string
            dfOWS['UI Job Code'] = dfOWS['UI Job Code'].astype(str)

            # Merge data
            self.result_dfOWS = self.filtered_dfOWS.merge(
                dfOWS,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfOWS.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfOWS.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged OWS data for pivot index.")

            # Create pivot table
            pivot_table = self.result_dfOWS.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.fillna(0).astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultOWSJobs = pivot_table
            self.styled_pivot_table_resultOWSJobs = self.pivot_table_resultOWSJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Find missing jobs
            self.missingjobsOWSresult = dfOWS[~dfOWS['UI Job Code'].isin(
                self.filtered_dfOWS['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsOWSresult.columns:
                self.missingjobsOWSresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsOWSresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultOWSJobs = pd.DataFrame({'Error': [f'OWS data processing failed: {str(e)}']})
            self.missingjobsOWSresult = pd.DataFrame()
