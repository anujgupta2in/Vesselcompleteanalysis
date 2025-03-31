import pandas as pd
import numpy as np

pd.set_option('future.no_silent_downcasting', True)

class BatterySystemProcessor:
    def __init__(self):
        self.filtered_dfbatteryjobs = pd.DataFrame()
        self.result_dfbattery = pd.DataFrame()
        self.pivot_table_resultbatteryJobs = pd.DataFrame()
        self.styled_pivot_table_resultbatteryJobs = None
        self.missingjobsbatteryresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_battery_data(self, df, dfbattery):
        try:
            self.filtered_dfbatteryjobs = df[df['Machinery Location'].str.contains('Battery', na=False)].copy()
            if 'Job Codecopy' not in self.filtered_dfbatteryjobs.columns:
                self.filtered_dfbatteryjobs['Job Codecopy'] = self.filtered_dfbatteryjobs['Job Code'].astype(str)
            self.filtered_dfbatteryjobs['Job Codecopy'] = self.filtered_dfbatteryjobs['Job Codecopy'].astype(object)
            self.filtered_dfbatteryjobs['Job Codecopy'] = self.filtered_dfbatteryjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dfbattery['UI Job Code'] = dfbattery['UI Job Code'].apply(self.safe_convert_to_string)

            self.result_dfbattery = self.filtered_dfbatteryjobs.merge(
                dfbattery,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dfbattery.reset_index(drop=True, inplace=True)

            title_col = next((col for col in ['Title', 'J3 Job Title', 'Task Description', 'Job Title'] if col in self.result_dfbattery.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged battery data for pivot index.")

            self.pivot_table_resultbatteryJobs = self.result_dfbattery.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            )
            self.pivot_table_resultbatteryJobs.replace(np.nan, '', inplace=True)
            self.pivot_table_resultbatteryJobs.replace('', -1, inplace=True)
            self.pivot_table_resultbatteryJobs = self.pivot_table_resultbatteryJobs.infer_objects(copy=False)
            self.pivot_table_resultbatteryJobs = self.pivot_table_resultbatteryJobs.astype(int)
            self.pivot_table_resultbatteryJobs = self.pivot_table_resultbatteryJobs.map(self.format_blank)

            self.styled_pivot_table_resultbatteryJobs = self.pivot_table_resultbatteryJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.missingjobsbatteryresult = dfbattery[~dfbattery['UI Job Code'].isin(
                self.filtered_dfbatteryjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobsbatteryresult.columns:
                self.missingjobsbatteryresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobsbatteryresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultbatteryJobs = pd.DataFrame({'Error': [f'Battery data processing failed: {str(e)}']})
            self.missingjobsbatteryresult = pd.DataFrame()
