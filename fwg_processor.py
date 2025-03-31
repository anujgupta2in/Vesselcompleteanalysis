import pandas as pd
import numpy as np

class FWGSystemProcessor:
    def __init__(self):
        self.filtered_dffwgjobs = pd.DataFrame()
        self.result_dffwg = pd.DataFrame()
        self.pivot_table_resultfwgJobs = pd.DataFrame()
        self.styled_pivot_table_resultfwgJobs = None
        self.missingjobsfwgresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_fwg_data(self, df, dffwg):
        try:
            fwg = ['Fresh Water Generator', 'Hydrophore System']

            self.filtered_dffwgjobs = df[df['Machinery Location'].str.contains('|'.join(fwg), na=False)].copy()

            if 'Job Codecopy' not in self.filtered_dffwgjobs.columns:
                self.filtered_dffwgjobs['Job Codecopy'] = self.filtered_dffwgjobs['Job Code'].astype(str)
            self.filtered_dffwgjobs['Job Codecopy'] = self.filtered_dffwgjobs['Job Codecopy'].astype(object)
            self.filtered_dffwgjobs['Job Codecopy'] = self.filtered_dffwgjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dffwg['UI Job Code'] = dffwg['UI Job Code'].astype(str)

            self.result_dffwg = self.filtered_dffwgjobs.merge(
                dffwg,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dffwg.reset_index(drop=True, inplace=True)

            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dffwg.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged FWG data for pivot index.")

            pivot_table = self.result_dffwg.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultfwgJobs = pivot_table
            self.styled_pivot_table_resultfwgJobs = self.pivot_table_resultfwgJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.missingjobsfwgresult = dffwg[~dffwg['UI Job Code'].isin(
                self.filtered_dffwgjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsfwgresult.columns:
                self.missingjobsfwgresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsfwgresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultfwgJobs = pd.DataFrame({'Error': [f'FWG data processing failed: {str(e)}']})
            self.missingjobsfwgresult = pd.DataFrame()
