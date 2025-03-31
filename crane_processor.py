import pandas as pd
import numpy as np

class CraneSystemProcessor:
    def __init__(self):
        self.filtered_dfcranejobs = pd.DataFrame()
        self.result_dfcrane = pd.DataFrame()
        self.pivot_table_resultcraneJobs = pd.DataFrame()
        self.styled_pivot_table_resultcraneJobs = None
        self.missingjobscraneresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_crane_data(self, df, dfcrane):
        try:
            # Define crane-related machinery locations
            crane_types = ['Crane', 'Bunker Davit']

            # Filter relevant rows
            self.filtered_dfcranejobs = df[df['Machinery Location'].str.contains('|'.join(crane_types), na=False)].copy()

            # Prepare 'Job Codecopy'
            if 'Job Codecopy' not in self.filtered_dfcranejobs.columns:
                self.filtered_dfcranejobs['Job Codecopy'] = self.filtered_dfcranejobs['Job Code'].astype(str)
            self.filtered_dfcranejobs['Job Codecopy'] = self.filtered_dfcranejobs['Job Codecopy'].astype(object)
            self.filtered_dfcranejobs['Job Codecopy'] = self.filtered_dfcranejobs['Job Codecopy'].apply(self.safe_convert_to_string)

            # Convert reference job codes
            dfcrane['UI Job Code'] = dfcrane['UI Job Code'].astype(str)

            # Merge filtered with reference
            self.result_dfcrane = self.filtered_dfcranejobs.merge(
                dfcrane,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfcrane.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfcrane.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged crane data for pivot index.")

            # Pivot table
            pivot_table = self.result_dfcrane.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultcraneJobs = pivot_table
            self.styled_pivot_table_resultcraneJobs = self.pivot_table_resultcraneJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Missing jobs
            self.missingjobscraneresult = dfcrane[~dfcrane['UI Job Code'].isin(
                self.filtered_dfcranejobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobscraneresult.columns:
                self.missingjobscraneresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobscraneresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultcraneJobs = pd.DataFrame({'Error': [f'Crane data processing failed: {str(e)}']})
            self.missingjobscraneresult = pd.DataFrame()