import pandas as pd
import numpy as np

class BoatSystemProcessor:
    def __init__(self):
        self.filtered_dfBoatjobs = pd.DataFrame()
        self.result_dfBoat = pd.DataFrame()
        self.pivot_table_resultBoatJobs = pd.DataFrame()
        self.styled_pivot_table_resultBoatJobs = None
        self.missingjobsBoatsresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_boat_data(self, df, dfBoats):
        try:
            # Define keywords to identify boats
            boats = ['Lifeboat','Lifeboat Davit', 'Liferaft/Rescue Boat Davit','Rescue Boat','Rescue Boat Davit','Boat','Liferaft']

            # Filter rows where 'Machinery Location' contains boat-related keywords
            self.filtered_dfBoatjobs = df[df['Machinery Location'].str.contains('|'.join(boats), na=False)].copy()

            # Ensure 'Job Codecopy' exists and is string
            if 'Job Codecopy' not in self.filtered_dfBoatjobs.columns:
                self.filtered_dfBoatjobs['Job Codecopy'] = self.filtered_dfBoatjobs['Job Code'].astype(str)
            self.filtered_dfBoatjobs['Job Codecopy'] = self.filtered_dfBoatjobs['Job Codecopy'].astype(object)
            self.filtered_dfBoatjobs['Job Codecopy'] = self.filtered_dfBoatjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            # Ensure reference codes are string
            dfBoats['UI Job Code'] = dfBoats['UI Job Code'].astype(str)

            # Merge data
            self.result_dfBoat = self.filtered_dfBoatjobs.merge(
                dfBoats,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfBoat.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfBoat.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged boat data for pivot index.")

            # Create pivot table
            pivot_table = self.result_dfBoat.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultBoatJobs = pivot_table
            self.styled_pivot_table_resultBoatJobs = self.pivot_table_resultBoatJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Find missing jobs
            self.missingjobsBoatsresult = dfBoats[~dfBoats['UI Job Code'].isin(
                self.filtered_dfBoatjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsBoatsresult.columns:
                self.missingjobsBoatsresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsBoatsresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultBoatJobs = pd.DataFrame({'Error': [f'Boat data processing failed: {str(e)}']})
            self.missingjobsBoatsresult = pd.DataFrame()
