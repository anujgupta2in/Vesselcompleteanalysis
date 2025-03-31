import pandas as pd
import numpy as np

class STPSystemProcessor:
    def __init__(self):
        self.filtered_dfSTP = pd.DataFrame()
        self.result_dfSTP = pd.DataFrame()
        self.pivot_table_resultSTPJobs = pd.DataFrame()
        self.styled_pivot_table_resultSTPJobs = None
        self.missingjobsSTPresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_stp_data(self, df, dfSTP):
        try:
            # Filter rows where 'Machinery Location' contains 'Sewage'
            self.filtered_dfSTP = df[df['Machinery Location'].str.contains('Sewage', na=False)].copy()

            # Ensure 'Job Codecopy' is of object type and safely converted to string
            if 'Job Codecopy' not in self.filtered_dfSTP.columns:
                self.filtered_dfSTP['Job Codecopy'] = self.filtered_dfSTP['Job Code'].astype(str)
            self.filtered_dfSTP['Job Codecopy'] = self.filtered_dfSTP['Job Codecopy'].astype(object)
            self.filtered_dfSTP['Job Codecopy'] = self.filtered_dfSTP['Job Codecopy'].apply(self.safe_convert_to_string)

            # Ensure reference codes are string
            dfSTP['UI Job Code'] = dfSTP['UI Job Code'].astype(str)

            # Merge data
            self.result_dfSTP = self.filtered_dfSTP.merge(
                dfSTP,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfSTP.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfSTP.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged STP data for pivot index.")

            # Create pivot table
            pivot_table = self.result_dfSTP.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.fillna(0).astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultSTPJobs = pivot_table
            self.styled_pivot_table_resultSTPJobs = self.pivot_table_resultSTPJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Find missing jobs
            self.missingjobsSTPresult = dfSTP[~dfSTP['UI Job Code'].isin(
                self.filtered_dfSTP['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsSTPresult.columns:
                self.missingjobsSTPresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsSTPresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultSTPJobs = pd.DataFrame({'Error': [f'STP data processing failed: {str(e)}']})
            self.missingjobsSTPresult = pd.DataFrame()
