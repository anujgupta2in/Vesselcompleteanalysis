import pandas as pd
import numpy as np

class BoilerSystemProcessor:
    def __init__(self):
        self.filtered_dfboilerjobs = pd.DataFrame()
        self.result_dfboiler = pd.DataFrame()
        self.pivot_table_resultboilerJobs = pd.DataFrame()
        self.styled_pivot_table_resultboilerJobs = None
        self.missingjobsboilerresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_boiler_data(self, df, dfboiler):
        try:
            boiler = ['Boiler']

            self.filtered_dfboilerjobs = df[df['Machinery Location'].str.contains('|'.join(boiler), na=False)].copy()

            if 'Job Codecopy' not in self.filtered_dfboilerjobs.columns:
                self.filtered_dfboilerjobs['Job Codecopy'] = self.filtered_dfboilerjobs['Job Code'].astype(str)
            self.filtered_dfboilerjobs['Job Codecopy'] = self.filtered_dfboilerjobs['Job Codecopy'].astype(object)
            self.filtered_dfboilerjobs['Job Codecopy'] = self.filtered_dfboilerjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dfboiler['UI Job Code'] = dfboiler['UI Job Code'].apply(self.safe_convert_to_string)

            self.result_dfboiler = self.filtered_dfboilerjobs.merge(
                dfboiler,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfboiler.reset_index(drop=True, inplace=True)

            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dfboiler.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged boiler data for pivot index.")

            pivot_table = self.result_dfboiler.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resultboilerJobs = pivot_table
            self.styled_pivot_table_resultboilerJobs = self.pivot_table_resultboilerJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.missingjobsboilerresult = dfboiler[~dfboiler['UI Job Code'].isin(
                self.filtered_dfboilerjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsboilerresult.columns:
                self.missingjobsboilerresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsboilerresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultboilerJobs = pd.DataFrame({'Error': [f'Boiler data processing failed: {str(e)}']})
            self.missingjobsboilerresult = pd.DataFrame()
