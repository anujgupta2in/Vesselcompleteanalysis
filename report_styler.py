class ReportStyler:
    """
    Class to provide styling for reports, particularly HTML reports.
    """
    
    def __init__(self):
        """Initialize the ReportStyler with default settings."""
        self.color_scheme = {
            'primary': '#1E88E5',
            'secondary': '#26A69A',
            'background': '#FFFFFF',
            'text': '#333333',
            'table_header': '#E3F2FD',
            'table_odd': '#F5F5F5',
            'table_even': '#FFFFFF',
            'table_border': '#DDDDDD',
            'link': '#0366D6',
            'visited_link': '#551A8B'
        }
        
        self.font_settings = {
            'family': 'Arial, sans-serif',
            'size_normal': '14px',
            'size_header': '18px',
            'size_title': '24px',
            'size_subtitle': '20px'
        }
        
        self.table_settings = {
            'width': '100%',
            'border_collapse': 'collapse',
            'cell_padding': '8px',
            'header_bg': self.color_scheme['table_header'],
            'row_hover': '#F1F8E9'
        }
    
    def get_color_scheme(self):
        """Get the current color scheme."""
        return self.color_scheme
    
    def get_font_settings(self):
        """Get the current font settings."""
        return self.font_settings
    
    def get_table_settings(self):
        """Get the current table settings."""
        return self.table_settings
    
    def generate_html_styles(self, colors=None, fonts=None, tables=None):
        """
        Generate CSS styling for HTML reports.
        
        Args:
            colors (dict, optional): Color scheme to use
            fonts (dict, optional): Font settings to use
            tables (dict, optional): Table settings to use
            
        Returns:
            str: CSS styles as a string
        """
        if colors is None:
            colors = self.color_scheme
        if fonts is None:
            fonts = self.font_settings
        if tables is None:
            tables = self.table_settings
            
        css = f"""
        /* Global Styles */
        body {{
            font-family: {fonts['family']};
            font-size: {fonts['size_normal']};
            line-height: 1.6;
            color: {colors['text']};
            background-color: {colors['background']};
            margin: 0;
            padding: 20px;
        }}
        
        /* Header Styles */
        h1, h2, h3, h4, h5, h6 {{
            color: {colors['primary']};
            margin-top: 20px;
            margin-bottom: 10px;
        }}
        
        h1 {{
            font-size: {fonts['size_title']};
            border-bottom: 2px solid {colors['primary']};
            padding-bottom: 10px;
            text-align: center;
        }}
        
        h2 {{
            font-size: {fonts['size_subtitle']};
            border-bottom: 1px solid {colors['secondary']};
            padding-bottom: 5px;
        }}
        
        h3 {{
            font-size: {fonts['size_header']};
        }}
        
        /* Table Styles */
        table {{
            width: {tables['width']};
            border-collapse: {tables['border_collapse']};
            margin-bottom: 20px;
            box-shadow: 0 2px 3px rgba(0, 0, 0, 0.1);
        }}
        
        th {{
            background-color: {tables['header_bg']};
            color: {colors['text']};
            font-weight: bold;
            padding: {tables['cell_padding']};
            text-align: left;
            border: 1px solid {colors['table_border']};
        }}
        
        td {{
            padding: {tables['cell_padding']};
            border: 1px solid {colors['table_border']};
        }}
        
        tr:nth-child(even) {{
            background-color: {colors['table_even']};
        }}
        
        tr:nth-child(odd) {{
            background-color: {colors['table_odd']};
        }}
        
        tr:hover {{
            background-color: {tables['row_hover']};
        }}
        
        /* Navigation Styles */
        .table-of-contents {{
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
        }}
        
        .table-of-contents ul {{
            list-style-type: none;
            padding-left: 20px;
        }}
        
        .table-of-contents li {{
            margin-bottom: 5px;
        }}
        
        a {{
            color: {colors['link']};
            text-decoration: none;
        }}
        
        a:hover {{
            text-decoration: underline;
        }}
        
        a:visited {{
            color: {colors['visited_link']};
        }}
        
        .table-nav {{
            text-align: right;
            margin-top: 10px;
            font-size: 0.9em;
        }}
        
        /* Section Styles */
        .section {{
            margin-bottom: 30px;
        }}
        
        .table-container {{
            margin-bottom: 20px;
            overflow-x: auto;
        }}
        
        /* Header and Footer Styles */
        .report-header {{
            margin-bottom: 30px;
            text-align: center;
        }}
        
        .vessel-info {{
            display: inline-block;
            text-align: left;
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin: 0 auto;
        }}
        
        .vessel-info p {{
            margin: 5px 0;
        }}
        
        .footer {{
            text-align: center;
            margin-top: 50px;
            padding-top: 10px;
            border-top: 1px solid {colors['table_border']};
            font-size: 0.9em;
            color: #666;
        }}
        
        /* Responsive Styles */
        @media screen and (max-width: 768px) {{
            body {{
                padding: 10px;
            }}
            
            table {{
                display: block;
                overflow-x: auto;
            }}
        }}
        """
        
        return css
    
    def style_dataframe(self, df):
        """
        Apply styling to a dataframe for display.
        
        Args:
            df (pandas.DataFrame): DataFrame to style
            
        Returns:
            pandas.io.formats.style.Styler: Styled DataFrame
        """
        return df.style.set_properties(**{
            'border': f'1px solid {self.color_scheme["table_border"]}',
            'padding': self.table_settings['cell_padding'],
            'text-align': 'left'
        }).set_table_styles([
            {'selector': 'th', 'props': [
                ('background-color', self.table_settings['header_bg']), 
                ('color', self.color_scheme['text']),
                ('font-weight', 'bold'),
                ('padding', self.table_settings['cell_padding']),
                ('border', f'1px solid {self.color_scheme["table_border"]}')
            ]},
            {'selector': 'tr:nth-child(even)', 'props': [
                ('background-color', self.color_scheme['table_even'])
            ]},
            {'selector': 'tr:nth-child(odd)', 'props': [
                ('background-color', self.color_scheme['table_odd'])
            ]}
        ])
