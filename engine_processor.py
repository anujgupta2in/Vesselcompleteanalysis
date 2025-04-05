import pandas as pd
import numpy as np
import re

def extract_units(job_data, unit_col):
    """Extract and sort unique units from the data."""
    try:
        return sorted(job_data[unit_col].dropna().unique(), key=lambda x: int(x) if str(x).isdigit() else x)
    except Exception:
        return []

def format_unit_data(unit_data):
    """Format unit data for display."""
    try:
        last_done_date = unit_data['Last Done Date'].dropna().astype(str).iloc[0] if not unit_data['Last Done Date'].dropna().empty else "No Date"
        last_done_running_hours = unit_data['Last Done Running Hours'].dropna().astype(str).iloc[0] if not unit_data['Last Done Running Hours'].dropna().empty else "No RH"
        remaining_hours = unit_data['Remaining Running Hours'].dropna().astype(str).iloc[0] if not unit_data['Remaining Running Hours'].dropna().empty else "No RH"
        return f"Date: {last_done_date}\nRH: {last_done_running_hours}\nRemaining Hours: {remaining_hours}"
    except Exception:
        return "Data formatting error"

def process_job_code_dynamic(data, job_codes, job_description, unit_pattern=r'Unit#(\d+)'):
    """Process job codes and return structured data."""
    if not isinstance(job_codes, list):
        job_codes = [job_codes]
    try:
        job_data = data[data['Job Code'].isin(job_codes)][
            ['Frequency', 'Last Done Date', 'Last Done Running Hours', 
             'Remaining Running Hours', 'Sub Component Location']
        ]
        job_data['Unit'] = job_data['Sub Component Location'].str.extract(unit_pattern)
        structured_data = pd.DataFrame({
            'Job Title': [job_description],
            'Frequency': [job_data['Frequency'].iloc[0] if not job_data.empty else "No Frequency"]
        })
        for unit in extract_units(job_data, 'Unit'):
            unit_data = job_data[job_data['Unit'] == unit]
            structured_data[f'Unit {unit}'] = [format_unit_data(unit_data)] if not unit_data.empty else ["No Data Available"]
        return structured_data
    except Exception as e:
        print(f"Error processing job code {job_codes}: {str(e)}")
        return pd.DataFrame({'Job Title': [job_description], 'Frequency': ["Error processing data"]})

def get_components_for_engine_type(engine_type):
    """Get the list of components to check based on engine type."""
    components = {
        "Normal Main Engine": [
            'Cylinder Liner - Main Engine',
            'Exhaust Valve - Main Engine',
            'Fuel Valve - Main Engine',
            'Start Air Valve - Main Engine',
            'Crosshead - Main Engine',
            'Turbocharger - Main Engine',
            'Main Bearing - Main Engine',
            'Auxiliary Blower - Main Engine',
            'Air Cooler - Main Engine',
            'Thrust Bearing - Main Engine',
            'Top Bracing - Main Engine',
            'Crankcase Relief Valve - Main Engine',
            'Exhaust Gas Receiver - Main Engine',
            'Oil Mist Detector - Main Engine'
        ],
        "MAN ME-C and ME-B Engine": [
            'Cylinder Liner - Main Engine',
            'Exhaust Valve - Main Engine',
            'Fuel Valve - Main Engine',
            'Start Air Valve - Main Engine',
            'Crosshead - Main Engine',
            'Turbocharger - Main Engine',
            'Main Bearing - Main Engine',
            'Auxiliary Blower - Main Engine',
            'Air Cooler - Main Engine',
            'Thrust Bearing - Main Engine',
            'Top Bracing - Main Engine',
            'EICU - Main Engine',
            'Axial Vibration Damper - Main Engine',
            'Bearing Condition Monitoring System - Main Engine',
            'Crankcase Relief Valve - Main Engine',
            'Exhaust Gas Receiver - Main Engine',
            'LDCL System - Main Engine',
            'Oil Mist Detector - Main Engine',
            'PMI Sensor - Main Engine',
            'RPM Pick Up Sensor - Main Engine',
            'Scavenge Receiver - Main Engine',
            'Stay Bolts - Main Engine',
            'Thrust Bearing - Main Engine',
            'Tie rod - Main Engine',
            'Torsion Vibration Damper - Main Engine',
            'Turning Gear - Main Engine',
            'Main Engine - HCU',
            'FIVA - Main Engine',
            'ELFI - Main Engine',
            'ELVA - Main Engine',
            'Exhaust Valve Actuator - Main Engine',
            'Fuel Pressure Booster - Main Engine',
            'Accumulator - Main Engine',
            'Hydraulic Start Up Pump - Main Engine'
        ],
        "RT Flex Engine": [
            'Cylinder Liner - Main Engine',
            'Exhaust Valve - Main Engine',
            'Fuel Valve - Main Engine',
            'Start Air Valve - Main Engine',
            'Crosshead - Main Engine',
            'Turbocharger - Main Engine',
            'Main Bearing - Main Engine',
            'Auxiliary Blower - Main Engine',
            'Air Cooler - Main Engine',
            'Thrust Bearing - Main Engine',
            'Top Bracing - Main Engine',
            'Crankcase Relief Valve - Main Engine',
            'Exhaust Gas Receiver - Main Engine',
            'Oil Mist Detector - Main Engine',
            'Turning Gear - Main Engine',
            'Cylinder Lubricator - Main Engine',
            'Connecting Rod - Main Engine',
            'Crankshaft - Main Engine',
            'Scavenge Receiver - Main Engine',
            'Camshaft - Main Engine',
            'Crankpin Bearing - Main Engine',
            'Angle Encoder - Main Engine',
            'Common Rail - Main Engine',
            'FCM - Main Engine',
            'Fuel Pump - Main Engine',
            'Fuel Pump Actuator - Main Engine',
            'Main Engine - Safety and Control',
            'Servo Oil Rail - Main Engine',
            'Servo Pump',
            'Servo Pump Drive - Main Engine',
            'WECS',
            'Exhaust Valve Control Unit - Main Engine',
            'Injection Control Unit - Main Engine'
        ],
        "RTA Engine": [
            'Cylinder Liner - Main Engine',
            'Exhaust Valve - Main Engine',
            'Fuel Valve - Main Engine',
            'Start Air Valve - Main Engine',
            'Crosshead - Main Engine',
            'Turbocharger - Main Engine',
            'Main Bearing - Main Engine',
            'Auxiliary Blower - Main Engine',
            'Air Cooler - Main Engine',
            'Thrust Bearing - Main Engine',
            'Top Bracing - Main Engine',
            'Crankcase Relief Valve - Main Engine',
            'Exhaust Gas Receiver - Main Engine',
            'Oil Mist Detector - Main Engine',
            'Turning Gear - Main Engine',
            'Fuel Injection Pump - Main Engine',
            'Cylinder Lubricator - Main Engine',
            'Connecting Rod - Main Engine',
            'Governor - Main Engine',
            'Crankshaft - Main Engine',
            'Scavenge Receiver - Main Engine',
            'Camshaft - Main Engine',
            'Crankpin Bearing - Main Engine'
        ],
        "UEC Engine": [
            'Cylinder Liner - Main Engine',
            'Exhaust Valve - Main Engine',
            'Fuel Valve - Main Engine',
            'Start Air Valve - Main Engine',
            'Crosshead - Main Engine',
            'Turbocharger - Main Engine',
            'Main Bearing - Main Engine',
            'Auxiliary Blower - Main Engine',
            'Air Cooler - Main Engine',
            'Thrust Bearing - Main Engine',
            'Top Bracing - Main Engine',
            'Advanced Electronic Cylinder Lubrication System',
            'Axial Vibration Damper - Main Engine',
            'Bearing Condition Monitoring System - Main Engine',
            'Crankcase Relief Valve - Main Engine',
            'Exhaust Gas Receiver - Main Engine',
            'LO Booster Pump - Advanced Electronic Cylinder Lubrication System',
            'Oil Mist Detector - Main Engine',
            'PMI Sensor - Main Engine',
            'RPM Pick Up Sensor - Main Engine',
            'Scavenge Receiver - Main Engine',
            'Stay Bolts - Main Engine',
            'Thrust Bearing - Main Engine',
            'Tie rod - Main Engine',
            'Torsion Vibration Damper - Main Engine',
            'Turning Gear - Main Engine',
            'Fuel Pressure Booster - Main Engine',
            'Electric Balancer - Main Engine',
            'Motor - LO Booster Pump',
            'Exhaust Valve Actuator - Main Engine',
            'Accumulator - Main Engine',
            'Hydraulic Start Up Pump - Main Engine',
            'Main Engine - HPS System',
            'Hydraulic Engine Driven Pump  - Main Engine',
            'Hydraulic Oil Filter - Main Engine',
            'Exhaust Valve HP Oil Pipe - Main Engine',
            'FO HP Pipe - Main Engine',
            'Engine Control Console - Main Engine',
            'Engine Electronic Control System - Main Engine',
            'Maneuvering System - Main Engine',
            'Pneumatic Control System - Main Engine',
            'Remote Control System - Main Engine',
            'Main Starting Air Valve - Main Engine',
            'Starting Air Distributor - Main Engine'
        ],
        "WINGD Engine": [
            'Cylinder Liner - Main Engine',
            'Exhaust Valve - Main Engine',
            'Fuel Valve - Main Engine',
            'Start Air Valve - Main Engine',
            'Crosshead - Main Engine',
            'Turbocharger - Main Engine',
            'Main Bearing - Main Engine',
            'Auxiliary Blower - Main Engine',
            'Air Cooler - Main Engine',
            'Thrust Bearing - Main Engine',
            'Top Bracing - Main Engine',
            'Crankcase Relief Valve - Main Engine',
            'Exhaust Gas Receiver - Main Engine',
            'Oil Mist Detector - Main Engine',
            'Turning Gear - Main Engine',
            'Cylinder Lubricator - Main Engine',
            'Connecting Rod - Main Engine',
            'Crankshaft - Main Engine',
            'Scavenge Receiver - Main Engine',
            'Camshaft - Main Engine',
            'Crankpin Bearing - Main Engine',
            'Angle Encoder - Main Engine',
            'Common Rail - Main Engine',
            'FCM - Main Engine',
            'Fuel Pump - Main Engine',
            'Fuel Pump Actuator - Main Engine',
            'Main Engine - Safety and Control',
            'Servo Oil Rail - Main Engine',
            'Servo Pump',
            'Servo Pump Drive - Main Engine',
            'WECS',
            'Exhaust Valve Control Unit - Main Engine',
            'Injection Control Unit - Main Engine'
        ]
    }
    return components.get(engine_type, components["Normal Main Engine"])

def process_engine_data(data, ref_sheet_path=None, engine_type=None):
    """Process both main and auxiliary engine data."""
    try:
        # Define job codes and descriptions for Main Engine
        job_code_data = [
            ([730, 805], "Stuffing Box Overhaul"),
            ([775, 776], "Piston Overhaul"),
            ([896], "Exhaust Valve Overhaul"),
            ([734], "Starting Air Overhaul"),
            ([860, 861, 862], "Fuel Valve Overhaul"),
            ([6795, 934], "Cylinder Liner Overhaul"),
            ([969, 5031, 802], "Main Bearing Overhaul - Main Engine", r'Main Bearing - Main Engine#(\d+)'),
            ([715], "Turbocharger Overhaul - Main Engine", r'Turbocharger - Main Engine#(\d+)'),
            ([873], "Fuel Injection Pump Overhaul - Main Engine", r'Fuel Injection Pump - Main Engine#(\d+)'),
            ([880], "FO Pressure Booster Overhaul - Main Engine", r'Main Engine - HCU#(\d+)'),
            ([903], "ELFI Overhaul - Main Engine", r'Main Engine - HCU#(\d+)'),
            ([901], "ELVA Overhaul - Main Engine", r'Main Engine - HCU#(\d+)'),
            ([885], "FIVA Overhaul - Main Engine", r'Main Engine - HCU#(\d+)')
        ]

        # Process Main Engine Data
        main_engine_data = []
        for job_info in job_code_data:
            if len(job_info) == 3:
                job_codes, job_description, unit_pattern = job_info
                structured_data = process_job_code_dynamic(data, job_codes, job_description, unit_pattern)
            else:
                job_codes, job_description = job_info
                structured_data = process_job_code_dynamic(data, job_codes, job_description)
            main_engine_data.append(structured_data)

        # Combine processed data
        main_engine_data = pd.concat(main_engine_data, ignore_index=True)

        # Extract Cylinder Unit information
        main_engine_filtered = data[data['Machinery Location'].str.contains("Main Engine", na=False)].copy()
        main_engine_filtered.loc[:, 'Cylinder Unit'] = main_engine_filtered['Sub Component Location'].str.extract(r'(Cylinder Unit#\d+)')
        main_engine_filtered.loc[:, 'Sub Components'] = main_engine_filtered['Sub Component Location'].str.extract(r'Cylinder Unit#\d+ > (.*)')

        # Create cylinder unit pivot table
        cylinder_pivot_table = main_engine_filtered.pivot_table(
            index='Cylinder Unit',
            columns='Sub Components',
            values='Job Code',
            aggfunc='count',
            fill_value=0
        ).reset_index()
        cylinder_pivot_table.columns.name = None



        # Get running hours for Main Engine
        main_engine_rows = data[data['Machinery Location'].str.contains('Main Engine', case=False, na=False)]
        main_engine_running_hours = "Not Available"
        if 'Machinery Running Hours' in data.columns:
            running_hours = main_engine_rows['Machinery Running Hours'].dropna()
            if not running_hours.empty:
                main_engine_running_hours = str(int(float(running_hours.iloc[0])))

        # Get auxiliary engine data
        aux_engine_data = data[data['Machinery Location'].str.contains("Auxiliary Engine", na=False, case=False)].copy()
        aux_running_hours = {'AE1': "Not Available", 'AE2': "Not Available", 'AE3': "Not Available"}
        if 'Machinery Running Hours' in data.columns:
            for ae_num in range(1, 4):
                ae_data = data[data['Machinery Location'].str.contains(f'Auxiliary Engine#{ae_num}', case=False, na=False)]
                if not ae_data.empty:
                    running_hours = ae_data['Machinery Running Hours'].dropna()
                    if not running_hours.empty:
                        aux_running_hours[f'AE{ae_num}'] = str(int(float(running_hours.iloc[0])))

        # Create standard pivot table
        pivot_table = main_engine_data.pivot_table(
            index='Job Title',
            values='Frequency',
            aggfunc='first'
        ).reset_index()

        # Process reference sheet if provided
        ref_pivot_table = None
        missing_jobs = None
        if ref_sheet_path is not None and engine_type is not None:
            try:
                sheet_mapping = {
                    "Normal Main Engine": "ME Jobs",
                    "MAN ME-C and ME-B Engine": "MEMEC",
                    "RT Flex Engine": "MERTFLEX",
                    "RTA Engine": "MERTA",
                    "UEC Engine": "MEUEC",
                    "WINGD Engine": "MEWINGD"
                }
                sheet_name = sheet_mapping.get(engine_type, "ME Jobs")

                ref_df = pd.read_excel(ref_sheet_path, sheet_name=sheet_name)
                ref_df['UI Job Code'] = ref_df['UI Job Code'].astype(str)

                filtered_dfMEjobs = data[data['Machinery Location'].str.contains('Main Engine', na=False)].copy()
                filtered_dfMEjobs['Job Codecopy'] = filtered_dfMEjobs['Job Code'].astype(str)

                result_dfME = filtered_dfMEjobs.merge(ref_df, left_on='Job Codecopy', right_on='UI Job Code', suffixes=('_filtered', '_ref'))

                ref_pivot_table = result_dfME.pivot_table(
                    index='Title',
                    columns='Machinery Location',
                    values='Job Codecopy',
                    aggfunc='count',
                    fill_value=0
                ).reset_index()
                ref_pivot_table.columns.name = None

                missing_jobs = ref_df[~ref_df['UI Job Code'].isin(filtered_dfMEjobs['Job Codecopy'])]
                if 'Remarks' in missing_jobs.columns:
                    missing_jobs = missing_jobs.drop(columns=['Remarks'])
                missing_jobs.reset_index(drop=True, inplace=True)
            except Exception as e:
                print(f"Error processing reference sheet: {e}")

        # Analyze engine components if engine type is provided
        component_status = None
        missing_count = 0
        if engine_type:
            components_to_check = get_components_for_engine_type(engine_type)
            appended_set = set(data['Machinery Location'].fillna('')) | set(data['Sub Component Location'].fillna(''))

            component_list = []
            status_list = []
            for component in components_to_check:
                found = any(component in item for item in appended_set)
                component_list.append(component)
                status_list.append("Present" if found else "Missing")
                if not found:
                    missing_count += 1

            component_status = pd.DataFrame({
                'Component': component_list,
                'Status': status_list
            })

        return (main_engine_data, aux_engine_data, main_engine_running_hours, aux_running_hours,
                pivot_table, ref_pivot_table, missing_jobs, cylinder_pivot_table, None, component_status, missing_count)

    except Exception as e:
        print(f"Error in process_engine_data: {str(e)}")
        raise

def generate_html_report(vessel_name, main_engine_data, aux_engine_data, main_engine_running_hours, aux_running_hours, component_status, missing_count):
    """Generate HTML report with provided data."""
    try:
        html_template = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Main Engine and Auxiliary Engine Report</title>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    margin: 20px;
                    padding: 20px;
                    background-color: #f4f4f4;
                }}
                h1 {{
                    text-align: center;
                    color: #333;
                }}
                h2, h3 {{
                    color: #444;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    background-color: white;
                    margin-bottom: 20px;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: left;
                }}
                th {{
                    background-color: #007bff;
                    color: white;
                }}
                tr:nth-child(even) {{
                    background-color: #f2f2f2;
                }}
                .running-hours {{
                    background-color: white;
                    padding: 15px;
                    border-radius: 5px;
                    margin: 15px 0;
                }}
            </style>
        </head>
        <body>
            <h1>Main Engine and Auxiliary Engine Report</h1>
            <h2>Vessel Name: {vessel_name}</h2>
            <div class="running-hours">
                <h3>Main Engine Running Hours: {main_engine_running_hours}</h3>
                <h3>Auxiliary Engine Running Hours:</h3>
                <ul>
                    <li>Auxiliary Engine#1: {aux_running_hours['AE1']}</li>
                    <li>Auxiliary Engine#2: {aux_running_hours['AE2']}</li>
                    <li>Auxiliary Engine#3: {aux_running_hours['AE3']}</li>
                </ul>
            </div>
            <h2>Main Engine Maintenance Data</h2>
            {main_engine_data.to_html(index=False, classes='table table-striped')}
            <h2>Auxiliary Engine Maintenance Data</h2>
            {aux_engine_data.to_html(index=False, classes='table table-striped')}
            <h2>Main Engine Component Status</h2>
            {component_status.to_html(index=False, classes='table table-striped') if component_status is not None else "<p>No component status available.</p>"}
            <p>Missing Components: {missing_count}</p>

        </body>
        </html>
        """
        return html_template
    except Exception as e:
        print(f"Error generating HTML report: {str(e)}")
        raise
