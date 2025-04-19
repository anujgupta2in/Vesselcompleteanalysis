import pandas as pd
from machinery_analyzer import MachineryAnalyzer

# =============================
# Utility: Style Pivot Table
# =============================
def create_and_style_pivot_table(data, index_col, columns_col, values_col, fill_value=0, cmap='RdYlGn'):
    data = data.dropna(subset=[index_col])
    pivot_table = pd.pivot_table(
        data,
        index=index_col,
        columns=columns_col,
        values=values_col,
        aggfunc='count',
        fill_value=fill_value,
        margins=True,
        margins_name='Total'
    )
    all_unique_values = data[index_col].unique()
    pivot_table = pivot_table.reindex(all_unique_values, fill_value=0)
    pivot_table_sorted = pivot_table.loc[pivot_table.sum(axis=1).sort_values(ascending=True).index]

    def highlight_null_and_zero_counts(value):
        color = 'red' if pd.isnull(value) or value == 0 else 'green'
        return f'background-color: {color}'

    styled_pivot_table = pivot_table_sorted.style.background_gradient(cmap=cmap, axis=None).map(highlight_null_and_zero_counts)
    return styled_pivot_table


# =============================
# Utility: Highlight Unknown Source
# =============================
def highlight_unknown(row):
    return ['background-color: red' if val == 'Unknown' else '' for val in row]


# =============================
# Main Class: QuickViewAnalyzer
# =============================
class QuickViewAnalyzer:

    def __init__(self, df, dfML, dfCM, dfVSM):
        self.df = df.copy()
        self.dfML = dfML.copy()
        self.dfCM = dfCM.copy()
        self.dfVSM = dfVSM.copy()
        self.analyzer = MachineryAnalyzer()

        # Clean and normalize machinery locations
        self.df['Machinery Locationcopy'] = self.df['Machinery Location'].str.lower().str.strip()
        self.df['Machinery Location Clean'] = self.df['Machinery Locationcopy'].apply(self.analyzer.clean_machinery_location)

        self.dfML['Machinery Location'] = self.dfML['Machinery Location'].str.lower().str.strip()
        self.dfML['Machinery Location Clean'] = self.dfML['Machinery Location'].apply(self.analyzer.clean_machinery_location)

        self.dfCM['Critical Machinery'] = self.dfCM['Critical Machinery'].str.lower().str.strip()
        self.dfCM['Machinery Location Clean'] = self.dfCM['Critical Machinery'].apply(self.analyzer.clean_machinery_location)

        self.dfVSM['Vessel Specific Machinery'] = self.dfVSM['Vessel Specific Machinery'].str.lower().str.strip()
        self.dfVSM['Machinery Location Clean'] = self.dfVSM['Vessel Specific Machinery'].apply(self.analyzer.clean_machinery_location)

        self.vml_set = set(self.df['Machinery Location Clean'].dropna().str.lower())
        self.ml_set = set(self.dfML['Machinery Location Clean'].dropna().str.lower())
        self.cm_set = set(self.dfCM['Machinery Location Clean'].dropna().str.lower())
        self.vsm_set = set(self.dfVSM['Machinery Location Clean'].dropna().str.lower())

    def calculate_missing_and_diff(self):
        union_set = self.ml_set.union(self.cm_set, self.vsm_set)
        dif_machinery = self.vml_set - union_set
        missing_machinery = union_set - self.vml_set

        dif_df = pd.DataFrame({'Different Machinery on Vessel': list(dif_machinery)})
        dif_df['Different Machinery on Vessel'] = dif_df['Different Machinery on Vessel'].str.title()

        miss_df = pd.DataFrame({'Missing Machinery on Vessel': list(missing_machinery)})
        miss_df['Missing Machinery on Vessel'] = miss_df['Missing Machinery on Vessel'].str.title()

        return dif_df.sort_values(by='Different Machinery on Vessel'), miss_df.sort_values(by='Missing Machinery on Vessel')

    def generate_jobsource_summary(self):
        self.df['Job Source'] = self.df['Job Source'].fillna('Unknown')
        title_counts = self.df.groupby('Job Source')['Title'].count()
        total_series = pd.Series(title_counts.sum(), index=['Total'])
        title_counts_with_total = pd.concat([title_counts, total_series.rename('Title Count')])
        result = title_counts_with_total.reset_index()
        result.columns = ['Job Source', 'Title Count']
        return result.style.apply(highlight_unknown, subset=['Job Source']).hide(axis='index')

    def get_basic_counts(self, **missing_data_sources):
        # Get missing machinery count from comparison
        _, miss_df = self.calculate_missing_and_diff()
        self.missing_machinery_count = len(miss_df)

        # Vessel/job basics
        vesselname = self.df['Vessel'].iloc[0] if 'Vessel' in self.df.columns else 'Unknown'
        totaljobs = self.df['Title'].count()
        criticaljobscount = self.df['Unnamed: 0'].count() if 'Unnamed: 0' in self.df.columns else 0

        # Normalize onboard machinery
        available_machinery = self.df["Machinery Locationcopy"].dropna().astype(str).str.lower().str.strip().unique()
        missing_table_summary = []

        for label, df in missing_data_sources.items():
            if isinstance(df, pd.DataFrame):
                count = 0

                if "Machinery" in df.columns:
                    # Normalize job machinery
                    df["Machinery_normalized"] = df["Machinery"].astype(str).str.lower().str.strip()

                    # Filter only rows where machinery appears in onboard list
                    filtered_df = df[df["Machinery_normalized"].apply(
                        lambda x: any(x in onboard or onboard in x for onboard in available_machinery)
                    )]

                    count = len(filtered_df)
                else:
                    # If no Machinery column, assume global
                    count = len(df)

                # Always append result
                missing_table_summary.append({
                    "Machinery System": label.replace("_", " "),
                    "Missing Jobs Count": count
                })

        missing_jobs_df = pd.DataFrame(missing_table_summary).sort_values(by="Missing Jobs Count", ascending=False)
        self.missing_jobs_table = missing_jobs_df
        total_missing_jobs = missing_jobs_df['Missing Jobs Count'].sum() if not missing_jobs_df.empty else 0

        return vesselname, totaljobs, criticaljobscount, total_missing_jobs, missing_jobs_df, self.missing_machinery_count

def generate_jobsource_summary(self):
    self.df['Job Source'] = self.df['Job Source'].fillna('Unknown')
    title_counts = self.df.groupby('Job Source')['Title'].count()
    cms_code_count = self.df['CMS Code'].count()

    total_series = pd.Series(title_counts.sum(), index=['Total'])
    title_counts_with_total = pd.concat([title_counts, total_series.rename('Title Count')])
    result = title_counts_with_total.reset_index()
    result.columns = ['Job Source', 'Title Count']

    self.cms_code_count = cms_code_count  # 🔧 Stored for display in app.py
    return result.style.apply(highlight_unknown, subset=['Job Source']).hide(axis='index')
