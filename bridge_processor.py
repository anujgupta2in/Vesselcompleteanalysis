import pandas as pd
import numpy as np

class BridgeSystemProcessor:
    def __init__(self):
        self.filtered_dfbridgejobs = pd.DataFrame()
        self.result_dfbridge = pd.DataFrame()
        self.pivot_table_resultbridgeJobs = pd.DataFrame()
        self.styled_pivot_table_resultbridgeJobs = None
        self.missingjobsbridgeresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_bridge_data(self, df, dfbridge):
        try:
            # Define bridge equipment
            bridge_eqmt = ['Navigation Equipment', 'Search and Rescue', 'Communication Equipment']

            # Filter relevant rows by 'Function'
            self.filtered_dfbridgejobs = df[df['Function'].str.contains('|'.join(bridge_eqmt), na=False)].copy()

            # Prepare 'Job Codecopy'
            if 'Job Codecopy' not in self.filtered_dfbridgejobs.columns:
                self.filtered_dfbridgejobs['Job Codecopy'] = self.filtered_dfbridgejobs['Job Code'].astype(str)
            self.filtered_dfbridgejobs['Job Codecopy'] = self.filtered_dfbridgejobs['Job Codecopy'].astype(object)
            self.filtered_dfbridgejobs['Job Codecopy'] = self.filtered_dfbridgejobs['Job Codecopy'].apply(self.safe_convert_to_string)

            # Convert reference job codes
            dfbridge['UI Job Code'] = dfbridge['UI Job Code'].astype(str)

            # Merge filtered with reference
            self.result_dfbridge = self.filtered_dfbridgejobs.merge(
                dfbridge,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfbridge.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfbridge.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged bridge data for pivot index.")

            # Pivot table
            pivot_table = self.result_dfbridge.pivot_table(
                index=title_col,
                columns='Function',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultbridgeJobs = pivot_table
            self.styled_pivot_table_resultbridgeJobs = self.pivot_table_resultbridgeJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Missing jobs
            self.missingjobsbridgeresult = dfbridge[~dfbridge['UI Job Code'].isin(
                self.filtered_dfbridgejobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsbridgeresult.columns:
                self.missingjobsbridgeresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsbridgeresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultbridgeJobs = pd.DataFrame({'Error': [f'Bridge data processing failed: {str(e)}']})
            self.missingjobsbridgeresult = pd.DataFrame()
