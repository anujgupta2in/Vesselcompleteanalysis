import pandas as pd
import numpy as np

class TankSystemProcessor:
    def __init__(self):
        self.filtered_dftanksjobs = pd.DataFrame()
        self.result_dftanks = pd.DataFrame()
        self.pivot_table_resulttanksJobs = pd.DataFrame()
        self.styled_pivot_table_resulttanksJobs = None
        self.missingjobstankresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_tank_data(self, df, dftanks):
        try:
            tanks = [
                'Ballast System','Bilge and Sludge System','Fuel Oil Service System','Cargo Handling System',
                'Lubricating Oil Purification System','Fuel Oil Storage and Transfer System','Fresh Water System',
                'Cargo Ventilation System','Lubricating Oil Storage and Transfer System','Lubricating Oil Service System',
                'Steam and Condensate System','Cooling Fresh Water System','Fuel Oil Purification System',
                'Cooling Sea Water System','Stern Tube System','Waste Handling'
            ]

            self.filtered_dftanksjobs = df[df['Function'].str.contains('|'.join(tanks), na=False)].copy()

            if 'Job Codecopy' not in self.filtered_dftanksjobs.columns:
                self.filtered_dftanksjobs['Job Codecopy'] = self.filtered_dftanksjobs['Job Code'].astype(str)
            self.filtered_dftanksjobs['Job Codecopy'] = self.filtered_dftanksjobs['Job Codecopy'].astype(object)
            self.filtered_dftanksjobs['Job Codecopy'] = self.filtered_dftanksjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dftanks['UI Job Code'] = dftanks['UI Job Code'].astype(str)

            self.result_dftanks = self.filtered_dftanksjobs.merge(
                dftanks,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )
            self.result_dftanks.reset_index(drop=True, inplace=True)

            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dftanks.columns), None)
            if title_col is None:
                raise ValueError("No suitable title column found in merged tank data for pivot index.")

            pivot_table = self.result_dftanks.pivot_table(
                index=title_col,
                columns='Function',
                values='Job Codecopy',
                aggfunc='count'
            )
            pivot_table.replace(np.nan, '', inplace=True)
            pivot_table.replace('', -1, inplace=True)
            pivot_table = pivot_table.astype(int)
            pivot_table = pivot_table.applymap(self.format_blank)

            self.pivot_table_resulttanksJobs = pivot_table
            self.styled_pivot_table_resulttanksJobs = self.pivot_table_resulttanksJobs.style\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                ], overwrite=False)\
                .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")

            self.missingjobstankresult = dftanks[~dftanks['UI Job Code'].isin(
                self.filtered_dftanksjobs['Job Codecopy']
            )].copy()
            if 'Remarks' in self.missingjobstankresult.columns:
                self.missingjobstankresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobstankresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resulttanksJobs = pd.DataFrame({'Error': [f'Tank data processing failed: {str(e)}']})
            self.missingjobstankresult = pd.DataFrame()
