import pandas as pd

class ReportStyler:
    """A class to handle styling of dataframes for reports."""
    
    def __init__(self):
        """Initialize the ReportStyler with color mappings."""
        # Define color mappings for statuses
        self.status_colors = {
            'Present': '#28a745',  # Green
            'Missing': '#dc3545',  # Red
            'Different from Reference': '#ffc107',  # Amber
            'Missing from Vessel': '#17a2b8'  # Cyan
        }
        
        # Colors for highlighting standardized machinery names
        self.machinery_colors = {
            'Main Engine': '#d4f1f9',  # Light blue
            'Auxiliary Engine': '#d5f5e3',  # Light green
            'Steering Gear': '#fcf3cf',  # Light yellow
            'Pump': '#fae5d3',  # Light orange
            'Compressor': '#ebdef0',  # Light purple
        }
    
    def style_dataframe(self, df, standardized_column=None, status_column='Status'):
        """Apply styling to the DataFrame.
        
        Args:
            df: The DataFrame to style
            standardized_column: Column name containing standardized machinery names
            status_column: Column name containing status information
            
        Returns:
            Styled DataFrame
        """
        if df is None or df.empty:
            return df
        
        # Create a copy to prevent warnings
        styled_df = df.copy()
        
        # Define styling function for alternating row colors
        def highlight_rows(row):
            return ['background-color: #f5f5f5' if i % 2 == 0 else '' for i in range(len(row))]
        
        # Define styling function for standardized machinery names
        def highlight_standardized(row):
            if standardized_column and standardized_column in row:
                machinery_name = str(row[standardized_column])
                
                # Check if any key in machinery_colors is in the machinery name
                for key, color in self.machinery_colors.items():
                    if key in machinery_name:
                        return [f'background-color: {color}'] * len(row)
                        
                # Default light gray if no match
                return ['background-color: #f9f9f9'] * len(row)
            return [''] * len(row)
        
        # Define styling function for status
        def highlight_status(row):
            if status_column and status_column in row:
                status = row[status_column]
                color = self.status_colors.get(status, '')
                return [f'color: {color}; font-weight: bold' if col == status_column else '' for col in row.index]
            return [''] * len(row)
        
        # Apply styling
        if standardized_column and standardized_column in styled_df.columns:
            style_funcs = [highlight_standardized]
        else:
            style_funcs = [highlight_rows]
        
        if status_column and status_column in styled_df.columns:
            style_funcs.append(highlight_status)
        
        # Style the dataframe with all applicable functions
        for func in style_funcs:
            styled_df = styled_df.style.apply(func, axis=1)
        
        return styled_df
    
    def highlight_critical(self, df, critical_column):
        """Highlight critical machinery.
        
        Args:
            df: The DataFrame to style
            critical_column: Column name indicating if machinery is critical
            
        Returns:
            Styled DataFrame
        """
        if df is None or df.empty or critical_column not in df.columns:
            return df
        
        # Function to highlight critical rows
        def highlight_critical_rows(row):
            if row[critical_column]:
                return ['background-color: #ffcccb'] * len(row)  # Light red for critical
            return [''] * len(row)
        
        return df.style.apply(highlight_critical_rows, axis=1)
    
    def highlight_differences(self, df, column1, column2):
        """Highlight differences between two columns.
        
        Args:
            df: The DataFrame to style
            column1: First column to compare
            column2: Second column to compare
            
        Returns:
            Styled DataFrame
        """
        if df is None or df.empty or column1 not in df.columns or column2 not in df.columns:
            return df
        
        # Function to highlight differences
        def highlight_diff(row):
            styles = [''] * len(row)
            col1_idx = df.columns.get_loc(column1)
            col2_idx = df.columns.get_loc(column2)
            
            # Compare values and highlight if different
            if row[column1] != row[column2]:
                styles[col1_idx] = 'background-color: #ffcccb'  # Light red
                styles[col2_idx] = 'background-color: #ffcccb'  # Light red
            
            return styles
        
        return df.style.apply(highlight_diff, axis=1)