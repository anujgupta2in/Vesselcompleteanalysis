import streamlit as st

class ReportStyler:
    @staticmethod
    def get_color_scheme():
        """Get user-selected color scheme."""
        st.sidebar.subheader("Color Scheme")
        colors = {
            'primary': st.sidebar.color_picker("Primary Color", "#007bff"),
            'background': st.sidebar.color_picker("Background Color", "#f4f4f4"),
            'text': st.sidebar.color_picker("Text Color", "#333333"),
            'header_bg': st.sidebar.color_picker("Header Background", "#ffffff"),
            'table_header': st.sidebar.color_picker("Table Header", "#007bff"),
            'table_stripe': st.sidebar.color_picker("Table Stripe", "#f2f2f2")
        }
        return colors

    @staticmethod
    def get_font_settings():
        """Get user-selected font settings."""
        st.sidebar.subheader("Font Settings")
        fonts = {
            'main_font': st.sidebar.selectbox(
                "Main Font",
                ["Arial", "Helvetica", "Roboto", "Times New Roman"],
                index=0
            ),
            'header_size': st.sidebar.slider("Header Size (px)", 16, 32, 24),
            'text_size': st.sidebar.slider("Text Size (px)", 12, 24, 14)
        }
        return fonts

    @staticmethod
    def get_table_settings():
        """Get user-selected table settings."""
        st.sidebar.subheader("Table Settings")
        table_settings = {
            'striped': st.sidebar.checkbox("Striped Rows", True),
            'bordered': st.sidebar.checkbox("Bordered", True),
            'compact': st.sidebar.checkbox("Compact View", False),
            'hover': st.sidebar.checkbox("Hover Effect", True)
        }
        return table_settings

    def generate_html_styles(self, colors, fonts, table_settings):
        """Generate HTML CSS styles based on user settings."""
        css = f"""
            body {{
                font-family: {fonts['main_font']}, sans-serif;
                margin: 20px;
                padding: 20px;
                background-color: {colors['background']};
                color: {colors['text']};
                font-size: {fonts['text_size']}px;
            }}
            h1 {{
                text-align: center;
                color: {colors['text']};
                font-size: {fonts['header_size']}px;
            }}
            h2, h3 {{
                color: {colors['text']};
                font-size: {int(fonts['header_size'] * 0.8)}px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                background-color: {colors['header_bg']};
                margin-bottom: 20px;
            }}
            th, td {{
                padding: {8 if not table_settings['compact'] else 4}px;
                text-align: left;
                border: {1 if table_settings['bordered'] else 0}px solid #ddd;
            }}
            th {{
                background-color: {colors['table_header']};
                color: white;
            }}
            tr:nth-child(even) {{
                background-color: {colors['table_stripe'] if table_settings['striped'] else 'transparent'};
            }}
            {f"tr:hover {{background-color: {colors['table_stripe']}}}" if table_settings['hover'] else ""}
            .running-hours {{
                background-color: {colors['header_bg']};
                padding: 15px;
                border-radius: 5px;
                margin: 15px 0;
            }}
        """
        return css

    def apply_excel_styling(self, writer, colors, table_settings):
        """Apply styling to Excel workbook."""
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': colors['table_header'],
            'font_color': 'white',
            'border': 1 if table_settings['bordered'] else 0
        })
        
        stripe_format = workbook.add_format({
            'bg_color': colors['table_stripe'],
            'border': 1 if table_settings['bordered'] else 0
        })
        
        regular_format = workbook.add_format({
            'border': 1 if table_settings['bordered'] else 0
        })
        
        return header_format, stripe_format, regular_format
