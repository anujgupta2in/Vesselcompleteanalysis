import pandas as pd
import numpy as np

class HPSCRSystemProcessor:
    def __init__(self):
        self.filtered_dfHPSCRjobs = pd.DataFrame()
        self.result_dfHPSCR = pd.DataFrame()
        self.pivot_table_resultHPSCRJobs = pd.DataFrame()
        self.styled_pivot_table_resultHPSCRJobs = None
        self.missingjobsHPSCRresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_hpscr_data(self, df, dfHPSCR):
        try:
            HPSCR = ['HP SCR']
            self.filtered_dfHPSCRjobs = df[df['Machinery Location'].str.contains('|'.join(HPSCR), na=False)].copy()

            if 'Job Codecopy' not in self.filtered_dfHPSCRjobs.columns:
                self.filtered_dfHPSCRjobs['Job Codecopy'] = self.filtered_dfHPSCRjobs['Job Code'].astype(str)
            self.filtered_dfHPSCRjobs['Job Codecopy'] = self.filtered_dfHPSCRjobs['Job Codecopy'].astype(object)
            self.filtered_dfHPSCRjobs['Job Codecopy'] = self.filtered_dfHPSCRjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dfHPSCR['UI Job Code'] = dfHPSCR['UI Job Code'].astype(str)

            self.result_dfHPSCR = self.filtered_dfHPSCRjobs.merge(
                dfHPSCR,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfHPSCR.reset_index(drop=True, inplace=True)

            title_col = next((col for col in ['Title', 'J3 Job Title', 'Task Description', 'Job Title'] if col in self.result_dfHPSCR.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged HPSCR data for pivot index.")

            self.pivot_table_resultHPSCRJobs = self.result_dfHPSCR.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            self.pivot_table_resultHPSCRJobs.replace(np.nan, '', inplace=True)
            self.pivot_table_resultHPSCRJobs.replace('', -1, inplace=True)
            self.pivot_table_resultHPSCRJobs = self.pivot_table_resultHPSCRJobs.astype(int)
            self.pivot_table_resultHPSCRJobs = self.pivot_table_resultHPSCRJobs.map(self.format_blank)

            self.styled_pivot_table_resultHPSCRJobs = self.pivot_table_resultHPSCRJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.missingjobsHPSCRresult = dfHPSCR[~dfHPSCR['UI Job Code'].isin(
                self.filtered_dfHPSCRjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsHPSCRresult.columns:
                self.missingjobsHPSCRresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsHPSCRresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultHPSCRJobs = pd.DataFrame({'Error': [f'HPSCR data processing failed: {str(e)}']})
            self.missingjobsHPSCRresult = pd.DataFrame()
