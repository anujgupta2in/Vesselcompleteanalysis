import pandas as pd
import numpy as np

class MiscSystemProcessor:
    def __init__(self):
        self.result_dfmisc = pd.DataFrame()
        self.pivot_table_resultmiscJobs = pd.DataFrame()
        self.pivot_table_resultmiscJobstotal = pd.DataFrame()
        self.styled_pivot_table_resultmiscJobs = None
        self.styled_pivot_table_resultmiscJobstotal = None
        self.styled_missingmiscjobsresult = None
        self.missingmiscjobsresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_misc_data(self, df, dfmisc):
        try:
            dfcopy = df.copy()
            if 'Job Codecopy' not in dfcopy.columns:
                dfcopy['Job Codecopy'] = dfcopy['Job Code'].astype(str)
            dfcopy['Job Codecopy'] = dfcopy['Job Codecopy'].apply(self.safe_convert_to_string)
            dfmisc['UI Job Code'] = dfmisc['UI Job Code'].apply(self.safe_convert_to_string)

            self.result_dfmisc = dfcopy.merge(
                dfmisc,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfmisc.reset_index(drop=True, inplace=True)

            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfmisc.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged misc data for pivot index.")

            if 'Function' not in self.result_dfmisc.columns:
                raise ValueError("'Function' column is missing in the merged misc data.")

            self.pivot_table_resultmiscJobs = self.result_dfmisc.pivot_table(
                index=title_col,
                columns='Function',
                values='Job Codecopy',
                aggfunc='count'
            ).replace(np.nan, '', inplace=False).replace('', -1).astype(int).applymap(self.format_blank)

            self.styled_pivot_table_resultmiscJobs = self.pivot_table_resultmiscJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.pivot_table_resultmiscJobstotal = self.result_dfmisc.pivot_table(
                index=title_col,
                values='Job Codecopy',
                aggfunc='count'
            ).sort_values(by='Job Codecopy', ascending=False)

            self.pivot_table_resultmiscJobstotal = self.pivot_table_resultmiscJobstotal.replace(np.nan, '', inplace=False)
            self.pivot_table_resultmiscJobstotal = self.pivot_table_resultmiscJobstotal.replace('', -1)
            self.pivot_table_resultmiscJobstotal = self.pivot_table_resultmiscJobstotal.astype(int)
            self.pivot_table_resultmiscJobstotal = self.pivot_table_resultmiscJobstotal.applymap(self.format_blank)

            self.styled_pivot_table_resultmiscJobstotal = self.pivot_table_resultmiscJobstotal.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.missingmiscjobsresult = dfmisc[~dfmisc['UI Job Code'].isin(dfcopy['Job Codecopy'])].copy()
            if 'Remarks' in self.missingmiscjobsresult.columns:
                self.missingmiscjobsresult.drop(columns=['Remarks'], inplace=True)
            self.missingmiscjobsresult.reset_index(drop=True, inplace=True)

            self.styled_missingmiscjobsresult = self.missingmiscjobsresult.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultmiscJobs = pd.DataFrame({'Error': [f'Misc data processing failed: {str(e)}']})
            self.pivot_table_resultmiscJobstotal = pd.DataFrame()
            self.missingmiscjobsresult = pd.DataFrame()