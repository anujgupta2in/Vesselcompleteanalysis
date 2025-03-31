import pandas as pd
import numpy as np

class PumpSystemProcessor:
    def __init__(self):
        self.filtered_dfpump = pd.DataFrame()
        self.pivot_table_pump = pd.DataFrame()
        self.pivot_table_resultpumpJobs = pd.DataFrame()
        self.styled_pivot_table_pump = None
        self.styled_pivot_table_resultpumpJobs = None
        self.mappings = {'4406': '425', '428': '425', '2329': '425', '4656': '426', '7039': '426', '6001': '602'}

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def process_pump_data(self, df, dfpump):
        try:
            print("🚀 Starting pump data processing...")

            if 'UI Job Code' not in dfpump.columns:
                raise ValueError("Reference sheet missing 'UI Job Code' column")

            # Ensure Job Codecopy exists
            if 'Job Codecopy' not in df.columns:
                df['Job Codecopy'] = df['Job Code'].astype(str).str.strip()

            # 🔹 Pump Location Count Table
            pump_df = df[df['Machinery Location'].str.contains('pump', case=False, na=False)]
            self.pivot_table_pump = pump_df.pivot_table(index='Machinery Location', values='Title', aggfunc='count')

            # ✅ Red highlight for counts > 6
            self.styled_pivot_table_pump = self.pivot_table_pump.style.apply(
                lambda x: ['background-color: red; color: white;' if v > 6 else '' for v in x], axis=1
            ).set_table_styles([
                {'selector': 'th', 'props': [('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('text-align', 'center'), ('min-width', '150px')]},
                {'selector': 'td:first-child', 'props': [('text-align', 'left')]}
            ], overwrite=False)

            print("✅ Pump count table created")

            # 🔹 Filter and map job codes
            self.filtered_dfpump = df[df['Machinery Location'].str.contains('Pump', case=False, na=False)].copy()

            self.filtered_dfpump['Job Codecopy'] = self.filtered_dfpump['Job Code'].astype(str).str.strip()
            self.filtered_dfpump['Job Codecopy'] = self.filtered_dfpump['Job Codecopy'].apply(self.safe_convert_to_string)

            self.filtered_dfpump['Job Code'] = self.filtered_dfpump['Job Code'].astype(str)
            self.filtered_dfpump['Job Code'] = self.filtered_dfpump['Job Code'].map(self.mappings).fillna(self.filtered_dfpump['Job Code'])
            self.filtered_dfpump['Job Code'] = self.filtered_dfpump['Job Code'].astype(str)

            job_codespump = dfpump['UI Job Code'].astype(str).tolist()
            matching_jobspump = self.filtered_dfpump[self.filtered_dfpump['Job Code'].isin(job_codespump)]

            # 🔹 Create the result pivot table
            pivot_table_resultpumpJobs = matching_jobspump.pivot_table(
                index=['Machinery Location'],
                columns='Job Code',
                values='Title',
                aggfunc='count',
                margins=True,
                margins_name='Total'
            )

            if pivot_table_resultpumpJobs.empty:
                print("⚠️ No matching pump jobs found.")
                self.pivot_table_resultpumpJobs = pd.DataFrame()
                self.styled_pivot_table_resultpumpJobs = None
                return self.pivot_table_resultpumpJobs

            pivot_table_resultpumpJobs.fillna(0, inplace=True)
            pivot_table_resultpumpJobs = pivot_table_resultpumpJobs.astype(int)
            self.pivot_table_resultpumpJobs = pivot_table_resultpumpJobs.copy()

            print("✅ Mapped Job Code Summary pivot created")

            # ✅ Green highlight for non-zero job counts
            styled = self.pivot_table_resultpumpJobs.style.highlight_between(
                left=1, right=9999, props="background-color: #d4edda; color: black;"
            )

            styled = styled.set_table_styles([
                {'selector': 'th', 'props': [('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('text-align', 'center'), ('min-width', '200px')]},
                {'selector': 'td:first-child', 'props': [('text-align', 'left')]}
            ], overwrite=False)

            self.styled_pivot_table_resultpumpJobs = styled

            print("✅ Styling applied safely with highlight.")
            return self.pivot_table_resultpumpJobs

        except Exception as e:
            import traceback
            traceback.print_exc()
            return pd.DataFrame({'Error': [f'Pump data processing failed: {str(e)}']})
