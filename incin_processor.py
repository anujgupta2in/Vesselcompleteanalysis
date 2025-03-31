import pandas as pd
import numpy as np

class IncineratorSystemProcessor:
    def __init__(self):
        self.filtered_dfIncin = pd.DataFrame()
        self.result_dfIncin = pd.DataFrame()
        self.pivot_table_resultIncinJobs = pd.DataFrame()
        self.styled_pivot_table_resultIncinJobs = None
        self.missingjobsIncinresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_incin_data(self, df, dfIncin):
        try:
            # Filter rows where 'Machinery Location' contains 'Incinerator'
            self.filtered_dfIncin = df[df['Machinery Location'].str.contains('Incinerator', na=False)].copy()

            # Ensure 'Job Codecopy' is of object type and safely converted to string
            if 'Job Codecopy' not in self.filtered_dfIncin.columns:
                self.filtered_dfIncin['Job Codecopy'] = self.filtered_dfIncin['Job Code'].astype(str)
            self.filtered_dfIncin['Job Codecopy'] = self.filtered_dfIncin['Job Codecopy'].astype(object)
            self.filtered_dfIncin['Job Codecopy'] = self.filtered_dfIncin['Job Codecopy'].apply(self.safe_convert_to_string)

            # Ensure reference codes are string
            dfIncin['UI Job Code'] = dfIncin['UI Job Code'].astype(str)

            # Merge data
            self.result_dfIncin = self.filtered_dfIncin.merge(
                dfIncin,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfIncin.reset_index(drop=True, inplace=True)

            # Detect title column
            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfIncin.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged incinerator data for pivot index.")

            # Create pivot table
            pivot_table = self.result_dfIncin.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.fillna(0)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultIncinJobs = pivot_table
            self.styled_pivot_table_resultIncinJobs = self.pivot_table_resultIncinJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Find missing jobs
            self.missingjobsIncinresult = dfIncin[~dfIncin['UI Job Code'].isin(
                self.filtered_dfIncin['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsIncinresult.columns:
                self.missingjobsIncinresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsIncinresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultIncinJobs = pd.DataFrame({'Error': [f'Incinerator data processing failed: {str(e)}']})
            self.missingjobsIncinresult = pd.DataFrame()