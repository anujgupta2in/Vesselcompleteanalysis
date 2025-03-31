import pandas as pd
import numpy as np

class PowerDistSystemProcessor:
    def __init__(self):
        self.filtered_dfpowerdistjobs = pd.DataFrame()
        self.result_dfpowerdist = pd.DataFrame()
        self.pivot_table_resultpowerdistJobs = pd.DataFrame()
        self.styled_pivot_table_resultpowerdistJobs = None
        self.missingjobspowerdistresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_powerdist_data(self, df, dfpowerdist):
        try:
            # Define keywords for power distribution components
            powerdist = ['SwitchBoard', 'Transformer', 'Panel', 'Emergency Lighting', 'Lighting', 'Switchboard']

            # Filter the DataFrame for relevant machinery locations
            self.filtered_dfpowerdistjobs = df[df['Machinery Location'].str.contains('|'.join(powerdist), na=False)].copy()

            # Prepare Job Codecopy column
            if 'Job Codecopy' not in self.filtered_dfpowerdistjobs.columns:
                self.filtered_dfpowerdistjobs['Job Codecopy'] = self.filtered_dfpowerdistjobs['Job Code'].astype(str)
            self.filtered_dfpowerdistjobs['Job Codecopy'] = self.filtered_dfpowerdistjobs['Job Codecopy'].astype(object)
            self.filtered_dfpowerdistjobs['Job Codecopy'] = self.filtered_dfpowerdistjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            # Ensure UI Job Code is string
            dfpowerdist['UI Job Code'] = dfpowerdist['UI Job Code'].astype(str)

            # Merge with reference
            self.result_dfpowerdist = self.filtered_dfpowerdistjobs.merge(
                dfpowerdist,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfpowerdist.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfpowerdist.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged power distribution data for pivot index.")

            # Pivot table creation
            pivot_table = self.result_dfpowerdist.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultpowerdistJobs = pivot_table
            self.styled_pivot_table_resultpowerdistJobs = self.pivot_table_resultpowerdistJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Identify missing jobs
            self.missingjobspowerdistresult = dfpowerdist[~dfpowerdist['UI Job Code'].isin(
                self.filtered_dfpowerdistjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobspowerdistresult.columns:
                self.missingjobspowerdistresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobspowerdistresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultpowerdistJobs = pd.DataFrame({'Error': [f'Power Distribution data processing failed: {str(e)}']})
            self.missingjobspowerdistresult = pd.DataFrame()
