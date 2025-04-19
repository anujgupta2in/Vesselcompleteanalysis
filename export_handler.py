import pandas as pd
import io
import datetime
import base64
import matplotlib.pyplot as plt
from io import BytesIO

class ExportHandler:
    def __init__(self, data, engine_type):
        self.data = data
        self.engine_type = engine_type
        self.timestamp = datetime.datetime.now().strftime("%Y-%m-%d")
        self.vessel_name = data['Vessel'].iloc[0] if 'Vessel' in data.columns and not data.empty else "Vessel"

    def plot_to_base64(self, fig):
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches='tight')
        buf.seek(0)
        encoded = base64.b64encode(buf.read()).decode("utf-8")
        buf.close()
        return f'<img src="data:image/png;base64,{encoded}" style="max-width:100%;">'

    def export_all_tabs_to_html(self, tab_data_dict, totaljobs=0, total_missing_jobs=0,
                                total_machinery=0, missing_machinery=0, vesselname="Vessel",
                                criticaljobscount=0, main_engine_jobs=0, ae_jobs=0):
        from report_styler import ReportStyler

        styler = ReportStyler()
        css = styler.generate_html_styles(
            styler.get_color_scheme(),
            styler.get_font_settings(),
            styler.get_table_settings()
        )

        html = f"""
        <html>
        <head>
            <meta charset='UTF-8'>
            <style>
            {css}
            #homeButton {{
                position: fixed; bottom: 30px; right: 30px; z-index: 1000;
                background-color: #007bff; color: white; padding: 10px 15px;
                border: none; border-radius: 5px; text-decoration: none;
                font-size: 14px; box-shadow: 2px 2px 5px rgba(0,0,0,0.3);
            }}
            .metric-summary {{
                display: flex; flex-wrap: wrap; justify-content: space-around;
                margin-bottom: 20px;
            }}
            .metric-card {{
                background: #f8f9fa; padding: 20px; margin: 10px;
                border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                flex: 1 1 200px; text-align: center;
            }}
            .metric-title {{ font-weight: bold; color: #333; }}
            .metric-value {{ font-size: 1.5em; color: #007bff; }}
            </style>
        </head>
        <body>
        <a name="top"></a>
        <h1 style='text-align:center;'>Vessel Maintenance Report</h1>
        <hr><h2>Navigation</h2><ul>
        """

        for tab in tab_data_dict.keys():
            anchor = tab.replace(" ", "_").replace("(", "").replace(")", "")
            html += f'<li><a href="#{anchor}">{tab}</a></li>'
        html += "</ul><hr>"

        # Metrics block (top)
        html += f"""
        <div class="metric-summary">
            <div class="metric-card"><div class="metric-title">🛳️ Vessel</div><div class="metric-value">{vesselname}</div></div>
            <div class="metric-card"><div class="metric-title">🧾 Total Jobs</div><div class="metric-value">{totaljobs}</div></div>
            <div class="metric-card"><div class="metric-title">🚨 Critical Jobs</div><div class="metric-value">{criticaljobscount}</div></div>
            <div class="metric-card"><div class="metric-title">❌ Total Missing Jobs</div><div class="metric-value">{total_missing_jobs}</div></div>
            <div class="metric-card"><div class="metric-title">🔧 Missing Machinery</div><div class="metric-value">{missing_machinery}</div></div>
        </div>
        """

        for tab, tables in tab_data_dict.items():
            anchor = tab.replace(" ", "_").replace("(", "").replace(")", "")
            html += f'<h2 id="{anchor}">{tab}</h2>'

            if tab == "QuickView Summary":
                try:
                    missing_jobs_df = tables[0]

                    # PIE 1
                    fig1, ax1 = plt.subplots(figsize=(4, 4))
                    wedges1, texts1, autotexts1 = ax1.pie(
                        [totaljobs, total_missing_jobs],
                        labels=["Total Jobs", "Missing Jobs"],
                        autopct=lambda pct: f'{int(pct * (totaljobs + total_missing_jobs) / 100)} ({pct:.1f}%)'
                    )
                    ax1.axis("equal")
                    ax1.set_title("Total Jobs vs Missing Jobs")
                    html += self.plot_to_base64(fig1)
                    plt.close(fig1)

                    # PIE 2
                    fig2, ax2 = plt.subplots(figsize=(4, 4))
                    wedges2, texts2, autotexts2 = ax2.pie(
                        [total_machinery, missing_machinery],
                        labels=["Present", "Missing"],
                        autopct=lambda pct: f'{int(pct * (total_machinery + missing_machinery) / 100)} ({pct:.1f}%)'
                    )
                    ax2.axis("equal")
                    ax2.set_title("Machinery Summary")
                    html += self.plot_to_base64(fig2)
                    plt.close(fig2)

                    # BAR
                    fig3, ax3 = plt.subplots(figsize=(18, 6))
                    bars = ax3.bar(missing_jobs_df["Machinery System"], missing_jobs_df["Missing Jobs Count"], color='#5DADE2')

                    for bar in bars:
                        height = bar.get_height()
                        ax3.text(bar.get_x() + bar.get_width() / 2, height + 0.5, str(int(height)),
                                ha='center', va='bottom', fontsize=9)

                    ax3.set_title("Missing Jobs by Machinery System")
                    ax3.set_xlabel("Machinery System")
                    ax3.set_ylabel("Missing Jobs Count")
                    ax3.tick_params(axis='x', rotation=45, labelsize=9)  # Make label font smaller
                    ax3.grid(axis='y', linestyle='--', alpha=0.5)

                    fig3.tight_layout()  # Ensures everything fits within frame
                    html += self.plot_to_base64(fig3)
                    plt.close(fig3)

                except Exception as chart_err:
                    html += f"<p>Chart generation failed: {chart_err}</p>"

            for i, df in enumerate(tables):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    if "Main Engine" in tab:
                        titles = [
                            "Maintenance Data for Main Engine",
                            "Main Engine Cylinder Unit Analysis",
                            "Reference Analysis for Main Engine",
                            "Missing Jobs for Main Engine",
                            "Component Status Analysis for Main Engine",
                            "Number of Missing Components for Main Engine"
                        ]
                        title = titles[i] if i < len(titles) else f"Main Engine Table {i+1}"
                    elif "Auxiliary Engine" in tab:
                        titles = [
                            "Task Count Analysis for Auxiliary Engine",
                            "Component Distribution for Auxiliary Engine",
                            "Component Status Analysis for Auxiliary Engine",
                            "Reference Analysis for Auxiliary Engine",
                            "Missing Jobs for Auxiliary Engine"
                        ]
                        title = titles[i] if i < len(titles) else f"Auxiliary Engine Table {i+1}"
                    else:
                        title = f"Table {i+1}"

                    try:
                        def color_cells(val):
                            try:
                                val = int(val)
                                if val == 0: return "background-color: #dc3545"
                                elif val == 1: return "background-color: #28a745"
                                elif val > 5: return "background-color: #fd7e14"
                            except: return ""

                        styled_df = df.style.applymap(color_cells, subset=df.select_dtypes(include='number').columns)
                        html += f"<h4>{title}</h4>" + styled_df.to_html(index=False, border=0)
                    except Exception:
                        html += f"<h4>{title}</h4>" + df.to_html(index=False, border=0)

            html += '<a href="#top" id="homeButton">Home</a><hr>'

        html += "</body></html>"
        return html
