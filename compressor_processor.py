import pandas as pd
import numpy as np

class CompressorSystemProcessor:
    def __init__(self):
        self.filtered_dfCompressorjobs = pd.DataFrame()
        self.result_dfCompressor = pd.DataFrame()
        self.pivot_table_resultCompressorJobs = pd.DataFrame()
        self.styled_pivot_table_resultCompressorJobs = None
        self.missingjobsCompressorresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_compressor_data(self, df, dfCompressor):
        try:
            # Ensure Job Codecopy exists
            if 'Job Codecopy' not in df.columns:
                df['Job Codecopy'] = df['Job Code'].astype(str).str.strip()

            # Filter rows where 'Machinery Location' contains 'Compressor'
            self.filtered_dfCompressorjobs = df[df['Machinery Location'].str.contains('Compressor', na=False)].copy()

            # Ensure 'Job Codecopy' is of object type and safely converted to string
            self.filtered_dfCompressorjobs['Job Codecopy'] = self.filtered_dfCompressorjobs['Job Codecopy'].astype(object)
            self.filtered_dfCompressorjobs['Job Codecopy'] = self.filtered_dfCompressorjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            dfCompressor = dfCompressor.copy()
            dfCompressor['UI Job Code'] = dfCompressor['UI Job Code'].astype(str)

            # Rename column if needed
            if 'J3 Job Title' in dfCompressor.columns and 'Title' not in dfCompressor.columns:
                dfCompressor.rename(columns={'J3 Job Title': 'Title'}, inplace=True)

            # Merge to get full matched rows
            self.result_dfCompressor = self.filtered_dfCompressorjobs.merge(
                dfCompressor,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )

            self.result_dfCompressor.reset_index(drop=True, inplace=True)

            if not self.result_dfCompressor.empty:
                # Pivot Table Creation
                pivot_table = self.result_dfCompressor.pivot_table(
                    index='Title_filtered',
                    columns='Machinery Location',
                    values='Job Codecopy',
                    aggfunc='count'
                )

                pivot_table.replace(np.nan, '', inplace=True)
                pivot_table.replace('', -1, inplace=True)
                pivot_table = pivot_table.astype(int)
                pivot_table = pivot_table.applymap(self.format_blank)

                self.pivot_table_resultCompressorJobs = pivot_table
            else:
                self.pivot_table_resultCompressorJobs = pd.DataFrame()

            self.styled_pivot_table_resultCompressorJobs = self.pivot_table_resultCompressorJobs.style.set_table_styles([
    {'selector': 'th', 'props': [('font-weight', 'bold')]},
    {'selector': 'td', 'props': [('text-align', 'center'), ('min-width', '150px')]},
    {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
], overwrite=False)

            # Find missing jobs
            self.missingjobsCompressorresult = dfCompressor[~dfCompressor['UI Job Code'].isin(
                self.filtered_dfCompressorjobs['Job Codecopy']
            )].copy()

            if 'Remarks' in self.missingjobsCompressorresult.columns:
                self.missingjobsCompressorresult.drop(columns=['Remarks'], inplace=True)

            self.missingjobsCompressorresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resultCompressorJobs = pd.DataFrame({'Error': [f'Compressor data processing failed: {str(e)}']})
