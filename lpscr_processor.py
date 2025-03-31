import pandas as pd
import numpy as np

class LPSCRSystemProcessor:
    def __init__(self):
        self.filtered_dfLPSCRjobs = pd.DataFrame()
        self.result_dfLPSCR = pd.DataFrame()
        self.pivot_table_resultLPSCRJobs = pd.DataFrame()
        self.styled_pivot_table_resultLPSCRJobs = None
        self.missingjobsLPSCRresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_lpscr_data(self, df, dfLPSCR):
        try:
            # Filter LP SCR jobs
            LPSCR = ['LP SCR']
            self.filtered_dfLPSCRjobs = df[df['Machinery Location'].str.contains('|'.join(LPSCR), na=False)].copy()

            # Prepare Job Codecopy
            if 'Job Codecopy' not in self.filtered_dfLPSCRjobs.columns:
                self.filtered_dfLPSCRjobs['Job Codecopy'] = self.filtered_dfLPSCRjobs['Job Code'].astype(str)
            self.filtered_dfLPSCRjobs['Job Codecopy'] = self.filtered_dfLPSCRjobs['Job Codecopy'].astype(object)
            self.filtered_dfLPSCRjobs['Job Codecopy'] = self.filtered_dfLPSCRjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dfLPSCR['UI Job Code'] = dfLPSCR['UI Job Code'].astype(str)

            # Merge datasets
            self.result_dfLPSCR = self.filtered_dfLPSCRjobs.merge(
                dfLPSCR,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfLPSCR.reset_index(drop=True, inplace=True)

            # Pivot Table
            title_col = next((col for col in ['Title', 'J3 Job Title', 'Task Description', 'Job Title'] if col in self.result_dfLPSCR.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged LPSCR data for pivot index.")

            self.pivot_table_resultLPSCRJobs = self.result_dfLPSCR.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            self.pivot_table_resultLPSCRJobs.replace(np.nan, '', inplace=True)
            self.pivot_table_resultLPSCRJobs.replace('', -1, inplace=True)
            self.pivot_table_resultLPSCRJobs = self.pivot_table_resultLPSCRJobs.astype(int)
            self.pivot_table_resultLPSCRJobs = self.pivot_table_resultLPSCRJobs.map(self.format_blank)

            self.styled_pivot_table_resultLPSCRJobs = self.pivot_table_resultLPSCRJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Identify missing jobs
            self.missingjobsLPSCRresult = dfLPSCR[~dfLPSCR['UI Job Code'].isin(
                self.filtered_dfLPSCRjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsLPSCRresult.columns:
                self.missingjobsLPSCRresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsLPSCRresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultLPSCRJobs = pd.DataFrame({'Error': [f'LPSCR data processing failed: {str(e)}']})
            self.missingjobsLPSCRresult = pd.DataFrame()
