import pandas as pd
import numpy as np

class MooringSystemProcessor:
    def __init__(self):
        self.filtered_dfMooring = pd.DataFrame()
        self.result_dfMooring = pd.DataFrame()
        self.pivot_table_resultMooringJobs = pd.DataFrame()
        self.styled_pivot_table_resultMooringJobs = None
        self.missingjobsMooringresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_mooring_data(self, df, dfMooring):
        try:
            # Filter rows where 'Machinery Location' contains 'Anchor|Mooring|Chain|Tail'
            self.filtered_dfMooring = df[df['Machinery Location'].str.contains('Anchor|Mooring|Chain|Tail', na=False)].copy()

            # Ensure 'Job Codecopy' is of object type and safely converted to string
            if 'Job Codecopy' not in self.filtered_dfMooring.columns:
                self.filtered_dfMooring['Job Codecopy'] = self.filtered_dfMooring['Job Code'].astype(str)
            self.filtered_dfMooring['Job Codecopy'] = self.filtered_dfMooring['Job Codecopy'].astype(object)
            self.filtered_dfMooring['Job Codecopy'] = self.filtered_dfMooring['Job Codecopy'].apply(self.safe_convert_to_string)

            # Ensure reference codes are string
            dfMooring['UI Job Code'] = dfMooring['UI Job Code'].astype(str)

            # Merge data
            self.result_dfMooring = self.filtered_dfMooring.merge(
                dfMooring,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfMooring.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfMooring.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged mooring data for pivot index.")

            # Create pivot table
            pivot_table = self.result_dfMooring.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultMooringJobs = pivot_table
            self.styled_pivot_table_resultMooringJobs = self.pivot_table_resultMooringJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Find missing jobs
            self.missingjobsMooringresult = dfMooring[~dfMooring['UI Job Code'].isin(
                self.filtered_dfMooring['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsMooringresult.columns:
                self.missingjobsMooringresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsMooringresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultMooringJobs = pd.DataFrame({'Error': [f'Mooring data processing failed: {str(e)}']})
            self.missingjobsMooringresult = pd.DataFrame()
