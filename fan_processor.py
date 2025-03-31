import pandas as pd
import numpy as np

class FanSystemProcessor:
    def __init__(self):
        self.filtered_dffanjobs = pd.DataFrame()
        self.matching_jobsfan = pd.DataFrame()
        self.pivot_table_resultfanJobs = pd.DataFrame()
        self.pivot_table_fan = pd.DataFrame()
        self.styled_pivot_table_resultfanJobs = None
        self.styled_pivot_table_fan = None

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_fan_data(self, df, dffan):
        try:
            fan = ['Fan']

            self.filtered_dffanjobs = df[df['Machinery Location'].str.contains('|'.join(fan), na=False)].copy()

            if 'Job Codecopy' not in self.filtered_dffanjobs.columns:
                self.filtered_dffanjobs['Job Codecopy'] = self.filtered_dffanjobs['Job Code'].astype(str)
            self.filtered_dffanjobs['Job Codecopy'] = self.filtered_dffanjobs['Job Codecopy'].astype(object)
            self.filtered_dffanjobs['Job Codecopy'] = self.filtered_dffanjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dffan['UI Job Code'] = dffan['UI Job Code'].astype(str)

            self.matching_jobsfan = self.filtered_dffanjobs.merge(
                dffan,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.matching_jobsfan.reset_index(drop=True, inplace=True)

            # Pivot Table 1: Title-wise count by location
            self.pivot_table_resultfanJobs = self.matching_jobsfan.pivot_table(
                index=['Machinery Location', 'Sub Component Location'],
                columns='Title',
                values='Job Codecopy',
                aggfunc='count'
            ).replace(np.nan, '', inplace=False).replace('', -1).astype(int).applymap(self.format_blank)

            self.styled_pivot_table_resultfanJobs = self.pivot_table_resultfanJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            # Pivot Table 2: Count of Titles per location
            self.pivot_table_fan = self.filtered_dffanjobs.pivot_table(
                index=['Machinery Location', 'Sub Component Location'],
                values='Title',
                aggfunc='count'
            ).replace(np.nan, '', inplace=False).replace('', -1).astype(int).applymap(self.format_blank)

            self.styled_pivot_table_fan = self.pivot_table_fan.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultfanJobs = pd.DataFrame({'Error': [f'Fan data processing failed: {str(e)}']})
            self.pivot_table_fan = pd.DataFrame({'Error': [f'Fan pivot generation failed: {str(e)}']})
