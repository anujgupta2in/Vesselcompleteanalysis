import pandas as pd
import numpy as np

class SteeringSystemProcessor:
    def __init__(self):
        self.filtered_dfSteering = pd.DataFrame()
        self.result_dfSteering = pd.DataFrame()
        self.pivot_table_resultSteeringJobs = pd.DataFrame()
        self.styled_pivot_table_resultSteeringJobs = None
        self.missingjobsSteeringresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_steering_data(self, df, dfSteering):
        try:
            # Filter rows where 'Machinery Location' contains 'Steering' or 'Stern'
            self.filtered_dfSteering = df[df['Machinery Location'].str.contains('Steering|Stern', na=False)].copy()

            # Ensure 'Job Codecopy' is of object type and safely converted to string
            if 'Job Codecopy' not in self.filtered_dfSteering.columns:
                self.filtered_dfSteering['Job Codecopy'] = self.filtered_dfSteering['Job Code'].astype(str)
            self.filtered_dfSteering['Job Codecopy'] = self.filtered_dfSteering['Job Codecopy'].astype(object)
            self.filtered_dfSteering['Job Codecopy'] = self.filtered_dfSteering['Job Codecopy'].apply(self.safe_convert_to_string)

            # Ensure reference codes are string
            dfSteering['UI Job Code'] = dfSteering['UI Job Code'].astype(str)

            # Merge data
            self.result_dfSteering = self.filtered_dfSteering.merge(
                dfSteering,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfSteering.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfSteering.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged steering data for pivot index.")

            # Create pivot table
            pivot_table = self.result_dfSteering.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultSteeringJobs = pivot_table
            self.styled_pivot_table_resultSteeringJobs = self.pivot_table_resultSteeringJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Find missing jobs
            self.missingjobsSteeringresult = dfSteering[~dfSteering['UI Job Code'].isin(
                self.filtered_dfSteering['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsSteeringresult.columns:
                self.missingjobsSteeringresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsSteeringresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultSteeringJobs = pd.DataFrame({'Error': [f'Steering data processing failed: {str(e)}']})
            self.missingjobsSteeringresult = pd.DataFrame()
