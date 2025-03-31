import pandas as pd
import numpy as np

class WorkshopSystemProcessor:
    def __init__(self):
        self.filtered_dfworkshopjobs = pd.DataFrame()
        self.result_dfworkshop = pd.DataFrame()
        self.pivot_table_resultworkshopJobs = pd.DataFrame()
        self.styled_pivot_table_resultworkshopJobs = None
        self.missingjobsworkshopresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_workshop_data(self, df, dfworkshop):
        try:
            workshop = ['Workshop']

            self.filtered_dfworkshopjobs = df[df['Machinery Location'].str.contains('|'.join(workshop), na=False)].copy()

            if 'Job Codecopy' not in self.filtered_dfworkshopjobs.columns:
                self.filtered_dfworkshopjobs['Job Codecopy'] = self.filtered_dfworkshopjobs['Job Code'].astype(str)
            self.filtered_dfworkshopjobs['Job Codecopy'] = self.filtered_dfworkshopjobs['Job Codecopy'].astype(object)
            self.filtered_dfworkshopjobs['Job Codecopy'] = self.filtered_dfworkshopjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dfworkshop['UI Job Code'] = dfworkshop['UI Job Code'].apply(self.safe_convert_to_string)

            self.result_dfworkshop = self.filtered_dfworkshopjobs.merge(
                dfworkshop,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfworkshop.reset_index(drop=True, inplace=True)

            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfworkshop.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged workshop data for pivot index.")

            pivot_table = self.result_dfworkshop.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultworkshopJobs = pivot_table
            self.styled_pivot_table_resultworkshopJobs = self.pivot_table_resultworkshopJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.missingjobsworkshopresult = dfworkshop[~dfworkshop['UI Job Code'].isin(
                self.filtered_dfworkshopjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsworkshopresult.columns:
                self.missingjobsworkshopresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsworkshopresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultworkshopJobs = pd.DataFrame({'Error': [f'Workshop data processing failed: {str(e)}']})
            self.missingjobsworkshopresult = pd.DataFrame()
