import pandas as pd
import io
import datetime
import re

class ExportHandler:
    def __init__(self, data, engine_type):
        """Initialize the ExportHandler."""
        self.data = data
        self.engine_type = engine_type
        self.timestamp = datetime.datetime.now().strftime("%Y-%m-%d")
        self.vessel_name = data['Vessel'].iloc[0] if 'Vessel' in data.columns and not data.empty else "Vessel"
    
    def _apply_excel_styling(self, writer, df, sheet_name):
        """Apply styling to Excel worksheet."""
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Apply header format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        # Auto-adjust columns width
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            worksheet.set_column(col_idx, col_idx, column_width)
    
    def generate_full_report(self, main_engine_data, aux_engine_data, pivot_table, 
                           cylinder_pivot_table, component_status, missing_count,
                           main_engine_running_hours, aux_running_hours):
        """Generate a full report with all analysis."""
        try:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Summary Sheet
                summary_data = pd.DataFrame({
                    'Item': ['Vessel Name', 'Engine Type', 'Report Date', 
                             'Main Engine Running Hours', 'AE1 Running Hours', 
                             'AE2 Running Hours', 'AE3 Running Hours',
                             'Missing Components Count'],
                    'Value': [self.vessel_name, self.engine_type, self.timestamp,
                              main_engine_running_hours,
                              aux_running_hours['AE1'],
                              aux_running_hours['AE2'],
                              aux_running_hours['AE3'],
                              missing_count]
                })
                summary_data.to_excel(writer, sheet_name='Summary', index=False)
                self._apply_excel_styling(writer, summary_data, 'Summary')
                
                # Concatenate main engine data if it's a list of DataFrames
                if isinstance(main_engine_data, list):
                    me_data = pd.concat(main_engine_data)
                else:
                    me_data = main_engine_data
                
                # Main Engine Analysis
                if not me_data.empty:
                    me_data.to_excel(writer, sheet_name='Main Engine Analysis', index=False)
                    self._apply_excel_styling(writer, me_data, 'Main Engine Analysis')
                
                # Component Analysis
                if pivot_table is not None and not pivot_table.empty:
                    pivot_table.to_excel(writer, sheet_name='Component Analysis')
                    self._apply_excel_styling(writer, pivot_table, 'Component Analysis')
                
                # Cylinder Analysis
                if cylinder_pivot_table is not None and not cylinder_pivot_table.empty:
                    cylinder_pivot_table.to_excel(writer, sheet_name='Cylinder Analysis')
                    self._apply_excel_styling(writer, cylinder_pivot_table, 'Cylinder Analysis')
                
                # Component Status
                if component_status is not None and not component_status.empty:
                    component_status.to_excel(writer, sheet_name='Component Status', index=False)
                    self._apply_excel_styling(writer, component_status, 'Component Status')
                
                # Auxiliary Engine Analysis
                if aux_engine_data is not None and not aux_engine_data.empty:
                    aux_engine_data.to_excel(writer, sheet_name='Auxiliary Engine Data', index=False)
                    self._apply_excel_styling(writer, aux_engine_data, 'Auxiliary Engine Data')
                
            output.seek(0)
            return output.getvalue()
            
        except Exception as e:
            print(f"Error generating full report: {str(e)}")
            # Return a minimal report with error information
            output = io.BytesIO()
            pd.DataFrame({'Error': [f"Failed to generate report: {str(e)}"]}).to_excel(output, index=False)
            output.seek(0)
            return output.getvalue()
    
    def generate_main_engine_report(self, main_engine_data, pivot_table, 
                                  cylinder_pivot_table, component_status, missing_count,
                                  main_engine_running_hours):
        """Generate a report focused on main engine analysis."""
        try:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Summary Sheet
                summary_data = pd.DataFrame({
                    'Item': ['Vessel Name', 'Engine Type', 'Report Date', 
                             'Main Engine Running Hours', 'Missing Components Count'],
                    'Value': [self.vessel_name, self.engine_type, self.timestamp,
                              main_engine_running_hours, missing_count]
                })
                summary_data.to_excel(writer, sheet_name='Summary', index=False)
                self._apply_excel_styling(writer, summary_data, 'Summary')
                
                # Concatenate main engine data if it's a list of DataFrames
                if isinstance(main_engine_data, list):
                    me_data = pd.concat(main_engine_data)
                else:
                    me_data = main_engine_data
                
                # Main Engine Analysis
                if not me_data.empty:
                    me_data.to_excel(writer, sheet_name='Main Engine Analysis', index=False)
                    self._apply_excel_styling(writer, me_data, 'Main Engine Analysis')
                
                # Component Analysis
                if pivot_table is not None and not pivot_table.empty:
                    pivot_table.to_excel(writer, sheet_name='Component Analysis')
                    self._apply_excel_styling(writer, pivot_table, 'Component Analysis')
                
                # Cylinder Analysis
                if cylinder_pivot_table is not None and not cylinder_pivot_table.empty:
                    cylinder_pivot_table.to_excel(writer, sheet_name='Cylinder Analysis')
                    self._apply_excel_styling(writer, cylinder_pivot_table, 'Cylinder Analysis')
                
                # Component Status
                if component_status is not None and not component_status.empty:
                    component_status.to_excel(writer, sheet_name='Component Status', index=False)
                    self._apply_excel_styling(writer, component_status, 'Component Status')
                
            output.seek(0)
            return output.getvalue()
            
        except Exception as e:
            print(f"Error generating main engine report: {str(e)}")
            output = io.BytesIO()
            pd.DataFrame({'Error': [f"Failed to generate report: {str(e)}"]}).to_excel(output, index=False)
            output.seek(0)
            return output.getvalue()
    
    def generate_auxiliary_engine_report(self, aux_engine_data, aux_running_hours):
        """Generate a report focused on auxiliary engine analysis."""
        try:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Summary Sheet
                summary_data = pd.DataFrame({
                    'Item': ['Vessel Name', 'Report Date', 'AE1 Running Hours', 
                             'AE2 Running Hours', 'AE3 Running Hours'],
                    'Value': [self.vessel_name, self.timestamp,
                              aux_running_hours['AE1'],
                              aux_running_hours['AE2'],
                              aux_running_hours['AE3']]
                })
                summary_data.to_excel(writer, sheet_name='Summary', index=False)
                self._apply_excel_styling(writer, summary_data, 'Summary')
                
                # Auxiliary Engine Analysis
                if aux_engine_data is not None and not aux_engine_data.empty:
                    aux_engine_data.to_excel(writer, sheet_name='Auxiliary Engine Data', index=False)
                    self._apply_excel_styling(writer, aux_engine_data, 'Auxiliary Engine Data')
                    
                    # Task Analysis by Engine
                    ae_numbers = []
                    for loc in aux_engine_data['Machinery Location'].dropna().unique():
                        match = re.search(r'Auxiliary Engine[\s-]*#?(\d+)', loc)
                        if match:
                            ae_numbers.append(match.group(1))
                    
                    if ae_numbers:
                        task_analysis = pd.DataFrame({'AE Number': ae_numbers})
                        task_analysis['Task Count'] = task_analysis['AE Number'].apply(
                            lambda x: len(aux_engine_data[aux_engine_data['Machinery Location'].str.contains(
                                f'Auxiliary Engine[\s-]*#?{x}', regex=True, na=False)])
                        )
                        task_analysis.to_excel(writer, sheet_name='Task Analysis', index=False)
                        self._apply_excel_styling(writer, task_analysis, 'Task Analysis')
                
            output.seek(0)
            return output.getvalue()
            
        except Exception as e:
            print(f"Error generating auxiliary engine report: {str(e)}")
            output = io.BytesIO()
            pd.DataFrame({'Error': [f"Failed to generate report: {str(e)}"]}).to_excel(output, index=False)
            output.seek(0)
            return output.getvalue()