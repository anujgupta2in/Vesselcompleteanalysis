import streamlit as st
import pandas as pd
import numpy as np
from engine_processor import process_engine_data
from auxiliary_engine_processor import AuxiliaryEngineProcessor
from purifier_processor import PurifierProcessor
from bwts_processor import BWTSProcessor
from hatch_processor import HatchProcessor
from cargopumping_processor import CargoPumpingProcessor
from csv_validator import CSVValidator
from inertgas_processor import InertGasSystemProcessor
from cargohandling_processor import CargoHandlingSystemProcessor
from cargoventing_processor import CargoVentingSystemProcessor
from lsaffa_processor import LSAFFAProcessor
from ffasys_processor import FFASystemProcessor
from pump_processor import PumpSystemProcessor
from compressor_processor import CompressorSystemProcessor
from ladder_processor import LadderSystemProcessor
from boat_processor import BoatSystemProcessor
from mooring_processor import MooringSystemProcessor
from steering_processor import SteeringSystemProcessor
from incin_processor import IncineratorSystemProcessor
from stp_processor import STPSystemProcessor
from ows_processor import OWSSystemProcessor
from powerdist_processor import PowerDistSystemProcessor
from crane_processor import CraneSystemProcessor
from emg_processor import EmergencyGenSystemProcessor
from bridge_processor import BridgeSystemProcessor
from refac_processor import RefacSystemProcessor
from fan_processor import FanSystemProcessor
from tank_processor import TankSystemProcessor
from fwg_processor import FWGSystemProcessor
from workshop_processor import WorkshopSystemProcessor
from boiler_processor import BoilerSystemProcessor
from misc_processor import MiscSystemProcessor
from battery_processor import BatterySystemProcessor
from bt_processor import BTSystemProcessor
from lpscr_processor import LPSCRSystemProcessor
from hpscr_processor import HPSCRSystemProcessor
from lsamapping_processor import LSAMappingProcessor
from ffamapping_processor import FFAMappingProcessor
from inactive_processor import InactiveMappingProcessor
from criticaljobs_processor import CriticalJobsProcessor
from export_handler import ExportHandler
from machinery_analyzer import MachineryAnalyzer
from report_styler import ReportStyler
from quickview import QuickViewAnalyzer, create_and_style_pivot_table
 # Added import for ReportStyler


#


# === Processor Initializations ===

workshop_processor = WorkshopSystemProcessor()
fwg_processor = FWGSystemProcessor()
ae_processor = AuxiliaryEngineProcessor()
bwts_processor = BWTSProcessor()
compressor_processor = CompressorSystemProcessor()
ladder_processor = LadderSystemProcessor()
boat_processor = BoatSystemProcessor()
mooring_processor = MooringSystemProcessor()
hatch_processor = HatchProcessor()
steering_processor = SteeringSystemProcessor()
boiler_processor = BoilerSystemProcessor()
incin_processor = IncineratorSystemProcessor()
stp_processor = STPSystemProcessor()
ows_processor = OWSSystemProcessor()
powerdist_processor = PowerDistSystemProcessor()
crane_processor = CraneSystemProcessor()
emg_processor = EmergencyGenSystemProcessor()
bridge_processor = BridgeSystemProcessor()
refac_processor = RefacSystemProcessor()
lsamapping_processor = LSAMappingProcessor()
ffamapping_processor = FFAMappingProcessor()
tank_processor = TankSystemProcessor()
battery_processor = BatterySystemProcessor()
bt_processor = BTSystemProcessor()
cargohandling_processor = CargoHandlingSystemProcessor()
cargopumping_processor = CargoPumpingProcessor()
cargoventing_processor = CargoVentingSystemProcessor()
critical_processor = CriticalJobsProcessor()
emg_processor = EmergencyGenSystemProcessor()
hpscr_processor = HPSCRSystemProcessor()
inactive_processor = InactiveMappingProcessor()
inertgas_processor = InertGasSystemProcessor()
lpscr_processor = LPSCRSystemProcessor()
misc_processor = MiscSystemProcessor()
mooring_processor = MooringSystemProcessor()
powerdist_processor = PowerDistSystemProcessor()
purifier_processor = PurifierProcessor()
steering_processor = SteeringSystemProcessor()
stp_processor = STPSystemProcessor()
tank_processor = TankSystemProcessor()
workshop_processor = WorkshopSystemProcessor()

def color_binary_cells(val):
    try:
        val = int(val)
        if val == 0:
            return 'background-color: #f8d7da; color: #721c24'  # red
        elif val == 1:
            return 'background-color: #d4edda; color: #155724'  # green
        elif val > 5:
            return 'background-color: #ffeeba; color: #856404'  # orange/yellow
    except:
        pass
    return ''


# Configure page settings
st.set_page_config(page_title="Vessel Report", layout="wide")

# Add custom CSS for smooth scrolling and metric card styling
st.markdown("""
<style>
    html {
        scroll-behavior: smooth;
    }
    .metric-row {
        display: flex;
        justify-content: space-between;
        margin-bottom: 1rem;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.12);
        width: 100%;
        margin: 0 0.5rem;
    }
    .engine-type-info {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Session state for tab switching
if 'current_tab' not in st.session_state:
    st.session_state.current_tab = 0

st.title("Vessel Report")
st.header("Upload Data")

# File upload section
col1, col2 = st.columns(2)
with col1:
    file_type = st.radio("Select file type:", ["CSV", "Excel"])
    if file_type == "CSV":
        uploaded_file = st.file_uploader("Upload CSV file", type="csv", key="data_file")
    else:
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"], key="data_file")
with col2:
    st.subheader("Reference Sheet")
    ref_sheet = st.file_uploader("Upload Reference Sheet (Excel)", type=["xlsx"], key="ref_sheet")

if uploaded_file is not None:
    try:
        if file_type == "CSV":
            data = pd.read_csv(uploaded_file)
        else:
            data = pd.read_excel(uploaded_file)

        # Engine and Equipment Configuration section
        st.header("Engine and Equipment Configuration")

        # Engine Type Selection
        st.markdown("""
        <div class="engine-type-info">
        <h4>Please select your vessel's main engine type below:</h4>
        <p>The engine type selection helps in analyzing the correct maintenance tasks and components for your specific engine.</p>
        </div>
        """, unsafe_allow_html=True)

        engine_descriptions = {
            "Normal Main Engine": "Standard configuration main engine without electronic control",
            "MAN ME-C and ME-B Engine": "Electronic controlled MAN B&W ME-C and ME-B series engines",
            "RT Flex Engine": "Wärtsilä RT-flex series common rail engines",
            "RTA Engine": "Wärtsilä/Sulzer RTA series mechanical engines",
            "UEC Engine": "Mitsubishi UEC series engines",
            "WINGD Engine": "WinGD X series engines"
        }

        # Create radio buttons for engine selection with descriptions
        engine_type = st.radio(
            "Select Engine Type",
            list(engine_descriptions.keys()),
            index=1,
            help="Choose the type of main engine installed on your vessel"
        )

        # Show description for selected engine type
        st.info(engine_descriptions[engine_type])

        # BWTS Model Selection
        st.markdown("""
        <div class="engine-type-info">
        <h4>Select Ballast Water Treatment System (BWTS) Model:</h4>
        <p>The BWTS model selection helps in analyzing the correct reference data for your specific ballast water treatment system.</p>
        </div>
        """, unsafe_allow_html=True)

        bwts_models = {
            "BWTS Optimarine": "BWTSOpti",
            "BWTS Alfalaval": "BWTSAlfalaval",
            "BWTS Echlor": "BWTSEchlor",
            "BWTS ERMA": "BWTSERMA",
            "BWTS Sunrai": "BWTSSunrai",
            "BWTS Techcross": "BWTStechcross",
            "Other BWTS": "BWTS"
        }

        # Initialize BWTS model in session state if not already there
        if 'bwts_model' not in st.session_state:
            st.session_state.bwts_model = "BWTS Optimarine"

        # Create radio buttons for BWTS model selection
        bwts_model = st.radio(
            "Select BWTS Model",
            list(bwts_models.keys()),
            index=0,
            help="Choose the type of Ballast Water Treatment System installed on your vessel"
        )

        # Store the selected model in session state
        st.session_state.bwts_model = bwts_model

        # Display the selected sheet name that will be used
        st.success(f"Will use reference sheet: {bwts_models[bwts_model]} for BWTS analysis")

        st.header("Data Preview")
        st.dataframe(data.head(), use_container_width=True)
        st.subheader("Detected Columns")
        st.write("Found columns:", ", ".join(data.columns.tolist()))

        # Validate data
        validator = CSVValidator()
        is_valid, errors = validator.validate_data(data)

        # Check if auto-correction was applied
        if '_machinery_location_fixed' in data.columns:
            corrected_count = (data['Machinery Location'] != data['_machinery_location_fixed']).sum()
            if corrected_count > 0:
                # Use the corrected column instead
                data['Machinery Location'] = data['_machinery_location_fixed']
                st.info(f"Auto-corrected {corrected_count} machinery location entries (e.g., 'Auxiliary EngineNo4' → 'Auxiliary Engine#4')")
                # Remove the temporary column
                data = data.drop(columns=['_machinery_location_fixed'])

        # Show validation results
        if not is_valid:
            st.warning("Data Validation Issues Detected:")
            for error in errors:
                st.warning(error)
            st.info("Analysis will proceed with available valid data.")
        else:
            st.success("Data validation successful!")

        # Initialize AuxiliaryEngineProcessor
        ae_processor = AuxiliaryEngineProcessor()

        # Process main engine data
        if ref_sheet is None:
            st.warning("No reference sheet uploaded. Some analysis features will be limited.")
            (main_engine_data, aux_engine_data, main_engine_running_hours, aux_running_hours,
             pivot_table, ref_pivot_table, missing_jobs, cylinder_pivot_table, pivot_table_filteredAE,
             component_status, missing_count) = process_engine_data(data, engine_type=engine_type)
        else:
            (main_engine_data, aux_engine_data, main_engine_running_hours, aux_running_hours,
             pivot_table, ref_pivot_table, missing_jobs, cylinder_pivot_table, pivot_table_filteredAE,
             component_status, missing_count) = process_engine_data(data, ref_sheet, engine_type)

        # Get auxiliary engine data
        aux_engine_data = ae_processor.get_maintenance_data(data)
        aux_running_hours = ae_processor.extract_running_hours(data)

        vessel_name = data['Vessel'].iloc[0]
        st.header(f"Vessel: {vessel_name}")

        # Running Hours Summary
        st.subheader("Running Hours Summary")
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric("Main Engine", value=main_engine_running_hours if main_engine_running_hours != "Not Available" else "N/A",
                      help="Main Engine Running Hours")
        with m2:
            st.metric("AE1", value=aux_running_hours['AE1'] if aux_running_hours['AE1'] != "Not Available" else "N/A",
                      help="Auxiliary Engine 1 Running Hours")
        with m3:
            st.metric("AE2", value=aux_running_hours['AE2'] if aux_running_hours['AE2'] != "Not Available" else "N/A",
                      help="Auxiliary Engine 2 Running Hours")
        with m4:
            st.metric("AE3", value=aux_running_hours['AE3'] if aux_running_hours['AE3'] != "Not Available" else "N/A",
                      help="Auxiliary Engine 3 Running Hours")


        

        # Initialize Export Handler
        export_handler = ExportHandler(data, engine_type)



        # Add export buttons
        st.subheader("Export Options")
        col1, col2, col3 = st.columns(3)
     
        with col1:
            try:
                if st.button("📥 Export Full HTML Report", key="export_html_btn"):
                    export_handler = ExportHandler(data, engine_type)

                    # ✅ Re-run QuickViewAnalyzer to get missing_jobs_df
                    ref_sheets = pd.read_excel(ref_sheet, sheet_name=None)
                    dfML = ref_sheets.get('Machinery Location', pd.DataFrame())
                    dfCM = ref_sheets.get('Critical Machinery', pd.DataFrame())
                    dfVSM = ref_sheets.get('Vessel Specific Machinery', pd.DataFrame())

                    analyzer = QuickViewAnalyzer(data, dfML, dfCM, dfVSM)

                    ae_ref_pivot, ae_missing_jobs = ae_processor.process_reference_data(data, ref_sheet)
                    main_engine_data, *_ , missing_jobs, _ = process_engine_data(data, ref_sheet, engine_type)
                    aux_task_count = ae_processor.create_task_count_table(data)
                    aux_component_dist = ae_processor.create_component_distribution(data)
                    aux_component_status, aux_missing_component_count = ae_processor.analyze_components(data)
                    ref_sheets = pd.read_excel(ref_sheet, sheet_name=None)
                    dfML = ref_sheets.get('Machinery Location', pd.DataFrame())
                    dfCM = ref_sheets.get('Critical Machinery', pd.DataFrame())
                    dfVSM = ref_sheets.get('Vessel Specific Machinery', pd.DataFrame())

                    analyzer = QuickViewAnalyzer(data, dfML, dfCM, dfVSM)

                    ae_ref_pivot, ae_missing_jobs = ae_processor.process_reference_data(data, ref_sheet)
                    battery_processor.process_battery_data(data, ref_sheets.get("Battery", pd.DataFrame()))
                    boat_processor.process_boat_data(data, ref_sheets.get("Boats", pd.DataFrame()))
                    boiler_processor.process_boiler_data(data, ref_sheets.get("Boiler", pd.DataFrame()))
                    bridge_processor.process_bridge_data(data, ref_sheets.get("Bridge", pd.DataFrame()))
                    bt_processor.process_bt_data(data, ref_sheets.get("Bow Thruster", pd.DataFrame()))
                    bwts_missing_jobs = bwts_processor.process_reference_data(data, ref_sheet)

                    chs_processor = CargoHandlingSystemProcessor()
                    dfCargoHandling = pd.read_excel(ref_sheet, sheet_name="Cargohanding")
                    chs_processor.process_reference_data(data, ref_sheet)

                    cargopumping_processor.process_reference_data(data, ref_sheet)
                    cargoventing_processor.process_reference_data(data, ref_sheet)
                    compressor_processor.process_compressor_data(data, ref_sheets.get("Compressor", pd.DataFrame()))
                    crane_processor.process_crane_data(data, ref_sheets.get("Crane", pd.DataFrame()))
                    critical_processor.process_critical_data(data, ref_sheets.get("criticalmapping", pd.DataFrame()))
                    main_engine_data, *_ , missing_jobs, _ = process_engine_data(data, ref_sheet, engine_type)
                    ffamapping_processor.process_ffa_data(data, pd.read_excel(ref_sheet, sheet_name="ffamapping"))
                    fwg_processor.process_fwg_data(data, pd.read_excel(ref_sheet, sheet_name="FWG"))
                    missing_jobs_hatch = hatch_processor.process_reference_data(data, ref_sheet)
                    hpscr_processor.process_hpscr_data(data, pd.read_excel(ref_sheet, sheet_name="HPSCRHITACHI"))
                    inactive_processor.process_inactive_data(data, pd.read_excel(ref_sheet, sheet_name="inactivemapping"))
                    inertgas_processor.process_reference_data(data, ref_sheet)
                    ladder_processor.process_ladder_data(data, pd.read_excel(ref_sheet, sheet_name="Ladders"))
                    incin_processor.process_incin_data(data, pd.read_excel(ref_sheet, sheet_name="Incin"))
                    lpscr_processor.process_lpscr_data(data, pd.read_excel(ref_sheet, sheet_name="LPSCRYANMAR"))
                    lsamapping_processor.process_lsa_data(data, pd.read_excel(ref_sheet, sheet_name="lsamapping"))
                    misc_processor.process_misc_data(data, pd.read_excel(ref_sheet, sheet_name="Misc"))
                    mooring_processor.process_mooring_data(data, pd.read_excel(ref_sheet, sheet_name="Mooring"))
                    ows_processor.process_ows_data(data, pd.read_excel(ref_sheet, sheet_name="OWS"))
                    powerdist_processor.process_powerdist_data(data, pd.read_excel(ref_sheet, sheet_name="Powerdist"))
                    missingjobspurifierresult = purifier_processor.process_reference_data(data, ref_sheet)
                    refac_processor.process_refac_data(data, pd.read_excel(ref_sheet, sheet_name="Refac"))
                    steering_processor.process_steering_data(data, pd.read_excel(ref_sheet, sheet_name="Steering"))
                    stp_processor.process_stp_data(data, pd.read_excel(ref_sheet, sheet_name="STP"))
                    tank_processor.process_tank_data(data, pd.read_excel(ref_sheet, sheet_name="Tanks"))
                    workshop_processor.process_workshop_data(data, pd.read_excel(ref_sheet, sheet_name="Workshop"))


                    # Add more processors if needed here...

                    vesselname, totaljobs, criticaljobscount, total_missing_jobs, missing_jobs_df, missing_machinery_count = analyzer.get_basic_counts(
                                            ae_missing_jobs=ae_missing_jobs,
                                            battery_missing_jobs=battery_processor.missingjobsbatteryresult,
                                            boat_missing_jobs=boat_processor.missingjobsBoatsresult,
                                            boiler_missing_jobs=boiler_processor.missingjobsboilerresult,
                                            bridge_missing_jobs=bridge_processor.missingjobsbridgeresult,
                                            bt_missing_jobs=bt_processor.missingjobsBTresult,
                                            bwts_missing_jobs=bwts_missing_jobs,
                                            Cargo_Handling_System=chs_processor.missing_jobs_cargohandling.copy(),
                                            Cargo_Pumping_System=cargopumping_processor.missingjobscargopumpingresult.copy(),
                                            Cargo_Venting_System=cargoventing_processor.missing_jobs_cargovent.copy(),
                                            compressor_missing_jobs=compressor_processor.missingjobsCompressorresult,
                                            crane_missing_jobs=crane_processor.missingjobscraneresult,
                                            Critical_Jobs=critical_processor.missingcriticaljobsresult.copy(),
                                            Main_Engine=missing_jobs.copy(),
                                            FFA_Mapping=ffamapping_processor.missingffajobsresult.copy(),
                                            FWG_System=fwg_processor.missingjobsfwgresult.copy(),
                                            Hatch_System=missing_jobs_hatch.copy(),
                                            HPSCR_System=hpscr_processor.missingjobsHPSCRresult.copy(),
                                            Inactive_Jobs=inactive_processor.missinginactivejobsresult.copy(),
                                            Inert_Gas_System=inertgas_processor.missing_jobs_igsystem.copy(),
                                            Ladder_System=ladder_processor.missingjobsLadderresult.copy(),
                                            Incinerator_System=incin_processor.missingjobsIncinresult.copy(),
                                            LPSCR_System=lpscr_processor.missingjobsLPSCRresult.copy(),
                                            LSA_Mapping=lsamapping_processor.missinglsajobsresult.copy(),
                                            Misc_Jobs=misc_processor.missingmiscjobsresult.copy(),
                                            Mooring_System=mooring_processor.missingjobsMooringresult.copy(),
                                            OWS_System=ows_processor.missingjobsOWSresult.copy(),
                                            Power_Distribution_System=powerdist_processor.missingjobspowerdistresult.copy(),
                                            Purifier_System=missingjobspurifierresult.copy(),
                                            Refac_System=refac_processor.missingjobsrefacresult.copy(),
                                            Steering_System=steering_processor.missingjobsSteeringresult.copy(),
                                            STP_System=stp_processor.missingjobsSTPresult.copy(),
                                            Tank_System=tank_processor.missingjobstankresult.copy(),
                                            Workshop_System=workshop_processor.missingjobsworkshopresult.copy(),
                                        )


                    # ✅ Collect all exportable tables
                    all_tab_tables = {
                        "QuickView Summary": [missing_jobs_df],
                        "Main Engine": [
                            main_engine_data,                  # ➤ Maintenance Data for Main Engine
                            cylinder_pivot_table,             # ➤ Main Engine Cylinder Unit Analysis
                            ref_pivot_table,                  # ➤ Reference Analysis Main Engine
                            missing_jobs,                     # ➤ Missing Jobs for Main Engine
                            component_status,                 # ➤ Component Status Analysis for Main Engine
                            missing_count                     # ➤ Number of missing components for Main Engine
                        ],
                            "Auxiliary Engine": [
                                aux_task_count,
                                aux_component_dist,
                                aux_component_status,
                                ae_ref_pivot,
                                ae_missing_jobs
                            ],
                            
                    }

                    html_report = export_handler.export_all_tabs_to_html(
                                all_tab_tables,
                                totaljobs=totaljobs,
                                total_missing_jobs=total_missing_jobs,
                                total_machinery=len(dfML),
                                missing_machinery=missing_machinery_count,
                                vesselname=vesselname,
                                criticaljobscount=criticaljobscount
                            )


                    filename = f"{data['Vessel'].iloc[0]}_Maintenance_Report.html" if "Vessel" in data.columns else "Maintenance_Report.html"

                    st.success("✅ Full HTML Report generated successfully!")
                    st.download_button(
                        label="📄 Download Full Report",
                        data=html_report,
                        file_name=filename,
                        mime="text/html",
                        key="download_html_btn"
                    )

            except Exception as e:
                st.error(f"❌ Error exporting HTML report: {e}")


        with col2:
            if st.button("Export Main Engine Report"):
                me_export = export_handler.generate_main_engine_report(
                    main_engine_data, pivot_table, 
                    cylinder_pivot_table, component_status, missing_count,
                    main_engine_running_hours
                )
                st.success(f"Main Engine Report generated successfully!")
                st.download_button(
                    label="Download Main Engine Report",
                    data=me_export,
                    file_name=f"{vessel_name}_main_engine_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        with col3:
            if st.button("Export Auxiliary Engine Report"):
                ae_export = export_handler.generate_auxiliary_engine_report(
                    aux_engine_data, aux_running_hours
                )
                st.success(f"Auxiliary Engine Report generated successfully!")
                st.download_button(
                    label="Download Auxiliary Engine Report",
                    data=ae_export,
                    file_name=f"{vessel_name}_auxiliary_engine_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # Tabs for analysis
        def switch_tab(tab_index):
            st.session_state.current_tab = tab_index

        colA, colB, colC, colD, colE, colF, colG, colH, colI, colJ, colK = st.columns(11)



        with colA:
            if st.button("Main Engine Analysis", key="main_tab"):
                switch_tab(0)
        with colB:
            if st.button("Auxiliary Engine Analysis", key="aux_tab"):
                switch_tab(1)
        with colC:
            if st.button("Machinery Location Analysis", key="machinery_tab"):
                switch_tab(2)
        with colD:
            if st.button("Purifier Analysis", key="purifier_tab"):
                switch_tab(3)
        with colE:
            if st.button("BWTS Analysis", key="bwts_tab"):
                switch_tab(4)
        with colF:
            if st.button("Hatch Analysis", key="hatch_tab"):
                switch_tab(5)
        with colG:
            if st.button("Cargo Pumping Analysis", key="cargopumping_tab"):
                switch_tab(6)
        with colH:
            if st.button("Inert Gas System Analysis", key="ig_tab"):
                switch_tab(7)
        with colI:
            if st.button("Cargo Handling System Analysis", key="chs_tab"):
                switch_tab(8)
        with colJ:
            if st.button("Cargo Ventilation System Analysis", key="cvs_tab"):
                switch_tab(9)
        with colK:
            if st.button("LSA/FFA Analysis", key="lsaffa_tab"):
                switch_tab(10)
        # New row of tabs (e.g., below existing 10 tabs)
                # Second row of tabs (FFASYS + Pump Analysis)
        tab_row_2 = st.columns(11)

        with tab_row_2[0]:
            if st.button("Fire Fighting System Analysis", key="ffasys_tab"):
                switch_tab(11)
        with tab_row_2[1]:
            if st.button("Pump Analysis", key="pump_tab"):
                switch_tab(12)
        with tab_row_2[2]:
            if st.button("Compressor Analysis", key="compressor_tab"):
                switch_tab(13)
        with tab_row_2[3]:
            if st.button("Ladder Analysis", key="ladder_tab"):
                switch_tab(14)
        with tab_row_2[4]:
            if st.button("Boat Analysis", key="boat_tab"):
                switch_tab(15)
        with tab_row_2[5]:
            if st.button("Mooring Analysis", key="mooring_tab"):
                switch_tab(16)
        with tab_row_2[6]:
            if st.button("Steering Analysis", key="steering_tab"):
                switch_tab(17)
        with tab_row_2[7]:
            if st.button("Incinerator Analysis", key="incin_tab"):
                switch_tab(18)
        with tab_row_2[8]:
            if st.button("STP Analysis", key="stp_tab"):
                switch_tab(19)
        with tab_row_2[9]:
            if st.button("OWS Analysis", key="ows_tab"):
                switch_tab(20)
        with tab_row_2[10]:
            if st.button("Power Distribution Analysis", key="powerdist_tab"):
                switch_tab(21)

        tab_row_3 = st.columns(10)

        with tab_row_3[0]:
            if st.button("Crane Analysis", key="crane_tab"):
                switch_tab(22)
        with tab_row_3[1]:
            if st.button("Emergency Generator Analysis", key="emg_tab"):
                switch_tab(23)
        with tab_row_3[2]:
            if st.button("Bridge System Analysis", key="bridge_tab"):
                switch_tab(24)
        with tab_row_3[3]:
            if st.button("Reefer & AC Analysis", key="refac_tab"):
                switch_tab(25)
        with tab_row_3[4]:
            if st.button("Fan Analysis", key="fan_tab"):
                switch_tab(26)
        with tab_row_3[5]:
            if st.button("Tank System Analysis", key="tank_tab"):
                switch_tab(27)
        with tab_row_3[6]:
            if st.button("FWG & Hydrophore Analysis", key="fwg_tab"):
                switch_tab(28)
        with tab_row_3[7]:
            if st.button("Workshop Analysis", key="workshop_tab"):
                switch_tab(29)
        with tab_row_3[8]:
            if st.button("Boiler Analysis", key="boiler_tab"):
                switch_tab(30)
        with tab_row_3[9]:
            if st.button("Misc Analysis", key="misc_tab"):
                switch_tab(31)

        tab_row_4 = st.columns(8)

        with tab_row_4[0]:
            if st.button("Battery Analysis", key="battery_tab"):
                switch_tab(32)

        with tab_row_4[1]:
            if st.button("BT Analysis", key="bt_tab"):
                switch_tab(33)

        with tab_row_4[2]:
            if st.button("LPSCR Analysis", key="lpscr_tab"):
                switch_tab(34)

        with tab_row_4[3]:
            if st.button("HPSCR Analysis", key="hpscr_tab"):
                switch_tab(35)

        with tab_row_4[4]:
            if st.button("LSA Mapping", key="lsa_tab"):
                switch_tab(36)

        with tab_row_4[5]:
            if st.button("FFA Mapping", key="ffa_tab"):
                switch_tab(37)

        with tab_row_4[6]:
            if st.button("Inactive Mapping", key="inactive_tab"):
                switch_tab(38)

        with tab_row_4[7]:
            if st.button("Critical Mapping", key="critical_tab"):
                switch_tab(39)



            # Fifth row of tabs (QuickView only)
        tab_row_5 = st.columns(1)

        with tab_row_5[0]:
            if st.button("QuickView Summary", key="quickview_tab"):
                switch_tab(40)


        if st.session_state.current_tab == 0:
            st.header("Main Engine Analysis")

            # Add Cylinder Unit Analysis
            st.subheader("Main Engine Cylinder Unit Analysis")
            if cylinder_pivot_table is not None:
                styled_cylinder = cylinder_pivot_table.style.applymap(color_binary_cells)
                st.dataframe(styled_cylinder, use_container_width=True)

            if ref_sheet is not None and ref_pivot_table is not None:
                st.subheader("Reference Analysis Main Engine")
                styled_ref_pivot = ref_pivot_table.style.applymap(color_binary_cells)
                st.dataframe(styled_ref_pivot, use_container_width=True)
                st.subheader("Missing Jobs for Main Engine")
                st.dataframe(missing_jobs, use_container_width=True)
            st.subheader("Maintenance Data for Main Engine")
            st.dataframe(main_engine_data, use_container_width=True)
            # Component Status Analysis
            st.subheader("Component Status Analysis for Main Engine")
            if component_status is not None:
                def color_status(val):
                    color = '#28a745' if val == 'Present' else '#dc3545'
                    return f'color: {color}'

                styled_status = component_status.style.map(color_status, subset=['Status'])
                st.dataframe(styled_status, use_container_width=True)
                st.info(f"Number of missing components: {missing_count}")
        elif st.session_state.current_tab == 1:
            st.header("Auxiliary Engine Analysis")

            try:
                # Task Count Analysis
                st.subheader("Task Count Analysis for Auxiliary Engine")
                task_count = ae_processor.create_task_count_table(data)
                if task_count is not None and not task_count.empty:
                    # Add a safe display method that converts to HTML to avoid JS errors
                    st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                    # Also provide download option for this table
                    csv = task_count.to_csv(index=False)
                    st.download_button(
                        label="Download Task Count CSV",
                        data=csv,
                        file_name="auxiliary_engine_task_count.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No task count data available for auxiliary engines.")
            except Exception as e:
                st.error(f"Error displaying Task Count Analysis: {str(e)}")
                st.error(f"Error details: {type(e).__name__}")

            try:
                # Component Distribution
                st.subheader("Component Distribution for Auxiliary Engine")
                component_dist = ae_processor.create_component_distribution(data)
                if component_dist is not None and not component_dist.empty:
                        styled_component_distAE = component_dist.style.applymap(color_binary_cells)
                        st.dataframe(styled_component_distAE, use_container_width=True)
                else:
                    st.info("No component distribution data available for auxiliary engines.")
            except Exception as e:
                st.error(f"Error displaying Component Distribution: {str(e)}")

            try:
                # Component Status Analysis
                component_status, missing_count = ae_processor.analyze_components(data)
                if component_status is not None and not component_status.empty:
                    st.subheader("Component Status Analysis for Auxiliary Engine")
                    def color_status(val):
                        color = '#28a745' if val == 'Present' else '#dc3545'
                        return f'color: {color}'
                    styled_status = component_status.style.map(color_status, subset=['Status'])
                    st.dataframe(styled_status, use_container_width=True)
                    st.info(f"Number of missing components: {missing_count}")
                else:
                    st.info("No component status data available for auxiliary engines.")
            except Exception as e:
                st.error(f"Error displaying Component Status Analysis: {str(e)}")

            try:
                # Reference Analysis
                if ref_sheet is not None:
                    ae_ref_pivot, ae_missing_jobs = ae_processor.process_reference_data(data, ref_sheet)

                    # Always show reference analysis section heading
                    st.subheader("Reference Analysis for Auxiliary Engine")

                    if ae_ref_pivot is not None and not ae_ref_pivot.empty:
                        # Use the HTML rendering approach to avoid JavaScript errors
                        styled_ref_ae = ae_ref_pivot.style.applymap(color_binary_cells)
                        st.dataframe(styled_ref_ae, use_container_width=True)

                        # Provide download option for this reference analysis
                        csv = ae_ref_pivot.to_csv(index=False)
                        st.download_button(
                            label="Download Reference Analysis CSV",
                            data=csv,
                            file_name="auxiliary_engine_reference_analysis.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No reference analysis matching records found for auxiliary engines.")

                    # Always show missing jobs section heading    
                    st.subheader("Missing Jobs for Auxiliary Engine")

                    if ae_missing_jobs is not None and not ae_missing_jobs.empty:
                        # Use the HTML rendering approach to avoid JavaScript errors
                        st.write(ae_missing_jobs.to_html(index=False), unsafe_allow_html=True)

                        # Provide download option for missing jobs
                        csv = ae_missing_jobs.to_csv(index=False)
                        st.download_button(
                            label="Download Missing Jobs CSV",
                            data=csv,
                            file_name="auxiliary_engine_missing_jobs.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No missing jobs analysis available for auxiliary engines.")
            except Exception as e:
                st.error(f"Error displaying Reference Analysis: {str(e)}")
                st.error(f"Error details: {type(e).__name__}")

            try:
                # Maintenance Data
                st.subheader("Maintenance Data for Auxiliary Engine")
                if aux_engine_data is not None and not aux_engine_data.empty:
                    st.dataframe(aux_engine_data, use_container_width=True)
                else:
                    st.info("No maintenance data available for auxiliary engines.")
            except Exception as e:
                st.error(f"Error displaying Maintenance Data: {str(e)}")
        elif st.session_state.current_tab == 2:
            st.header("Machinery Location Analysis")

            # Initialize MachineryAnalyzer
            analyzer = MachineryAnalyzer()
            styler = ReportStyler() # Initialize ReportStyler

            if ref_sheet is not None:
                # Process data with reference sheet
                df_analyzed, analysis_results = analyzer.process_data(data, ref_sheet)

                # Add info about machinery location analysis
                st.header("Machinery Location Analysis")
                st.info("""
                The Machinery Location Analysis examines the equipment data and compares it against a reference sheet.
                Machinery locations are automatically standardized by removing suffixes, location indicators (port/starboard), 
                and number indicators to provide more accurate comparisons. Critical equipment is highlighted.
                """)

                # Display different machinery analysis
                st.subheader("Different Machinery on Vessel Analysis")
                if 'different_machinery' in analysis_results:
                    # Rename the column for better readability
                    different_df = analysis_results['different_machinery'].rename(columns={
                        'Machinery Location': 'Standardized Machinery Location'
                    })
                    st.dataframe(different_df, use_container_width=True, 
                                column_config={
                                    "Standardized Machinery Location": st.column_config.TextColumn(
                                        "Standardized Machinery Location",
                                        width="large",
                                    )
                                })

                # Display missing machinery analysis
                st.subheader("Missing Machinery on Vessel Analysis")
                if 'missing_machinery' in analysis_results:
                    # Rename the column for better readability
                    missing_df = analysis_results['missing_machinery'].rename(columns={
                        'Machinery Location': 'Standardized Machinery Location'
                    })
                    st.dataframe(missing_df, use_container_width=True,
                                column_config={
                                    "Standardized Machinery Location": st.column_config.TextColumn(
                                        "Standardized Machinery Location",
                                        width="large",
                                    )
                                })

                # These detailed machinery location analysis tables have been removed as requested
            else:
                st.warning("Reference sheet is required for machinery location analysis. Please upload a reference sheet.")

        elif st.session_state.current_tab == 3:
            st.header("Purifier Analysis")

            # Initialize PurifierProcessor
            pu_processor = PurifierProcessor()

            try:
                # Running Hours for Purifiers
                purifier_hours = pu_processor.extract_running_hours(data)

                if not purifier_hours.empty:
                    st.subheader("Purifier Running Hours")
                    st.dataframe(purifier_hours, use_container_width=True)

                    # Create download option for running hours
                    csv = purifier_hours.to_csv(index=False)
                    st.download_button(
                        label="Download Purifier Running Hours CSV",
                        data=csv,
                        file_name="purifier_running_hours.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No running hours data available for purifiers.")
            except Exception as e:
                st.error(f"Error displaying Purifier Running Hours: {str(e)}")

            try:
                # Task Count Analysis
                st.subheader("Task Count Analysis for Purifiers")
                task_count = pu_processor.create_task_count_table(data)
                if not task_count.empty:
                    # Use HTML rendering approach to avoid JS errors
                    st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                    # Provide download option
                    csv = task_count.to_csv(index=False)
                    st.download_button(
                        label="Download Task Count CSV",
                        data=csv,
                        file_name="purifier_task_count.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No task count data available for purifiers.")
            except Exception as e:
                st.error(f"Error displaying Task Count Analysis: {str(e)}")

            # Component Distribution and Component Analysis removed as requested

            try:
                # Reference Job Analysis
                if ref_sheet is not None:
                    # Get missing jobs data
                    missing_jobs_purifier = pu_processor.process_reference_data(data, ref_sheet)

                    # Generate result_dfpurifiers data
                    # Create a copy of data to avoid modifications to original
                    data_copy = data.copy()

                    # Filter for purifier jobs
                    filtered_dfpurifierjobs = data_copy[data_copy['Machinery Location'].str.contains('Purifier', case=False, na=False)].copy()

                    # Read the reference sheet
                    ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names

                    # Look for 'Purifiers' sheet specifically first
                    purifier_sheet = 'Purifiers' if 'Purifiers' in ref_sheet_names else None

                    # If not found, try to find any sheet with 'Purifier' or 'PU' in the name
                    if purifier_sheet is None:
                        for sheet in ref_sheet_names:
                            if 'Purifier' in sheet.lower() or 'PU' in sheet:
                                purifier_sheet = sheet
                                break

                    # If still no purifier-specific sheet, use the first sheet
                    if purifier_sheet is None:
                        purifier_sheet = ref_sheet_names[0]
                        print(f"No purifier sheet found in app.py, using the first sheet: {purifier_sheet}")
                    else:
                        print(f"Using reference sheet in app.py: {purifier_sheet}")

                    dfpurifiers = pd.read_excel(ref_sheet, sheet_name=purifier_sheet)

                    # Create Job Codecopy column
                    if 'Job Code' in filtered_dfpurifierjobs.columns:
                        filtered_dfpurifierjobs['Job Codecopy'] = filtered_dfpurifierjobs['Job Code'].astype(str)

                        # Display Reference Jobs for Purifiers
                        st.subheader("Reference Jobs for Purifiers")
                        if not filtered_dfpurifierjobs.empty and not dfpurifiers.empty:
                            try:
                                # Display the columns in the reference data (for debugging)
                                print(f"Columns in reference data: {dfpurifiers.columns.tolist()}")

                                # Check which column to use for job code in reference data
                                job_code_col = None
                                for possible_col in ['UI Job Code', 'Job Code', 'JobCode', 'Code']:
                                    if possible_col in dfpurifiers.columns:
                                        job_code_col = possible_col
                                        print(f"Found job code column: {job_code_col}")
                                        break

                                if job_code_col is None:
                                    st.error(f"No job code column found in reference data. Available columns: {', '.join(dfpurifiers.columns.tolist())}")

                                if job_code_col is not None:
                                    # Convert both columns to string to avoid type mismatch
                                    filtered_dfpurifierjobs['Job Codecopy'] = filtered_dfpurifierjobs['Job Codecopy'].astype(str)
                                    dfpurifiers[job_code_col] = dfpurifiers[job_code_col].astype(str)

                                    # Print sample values from both columns for debugging
                                    print(f"Sample Job Codecopy values: {filtered_dfpurifierjobs['Job Codecopy'].iloc[:5].tolist()}")
                                    print(f"Sample {job_code_col} values: {dfpurifiers[job_code_col].iloc[:5].tolist()}")

                                    # Merge filtered_dfpurifierjobs with dfpurifiers on matching job codes
                                    result_dfpurifiers = filtered_dfpurifierjobs.merge(
                                        dfpurifiers, 
                                        left_on='Job Codecopy', 
                                        right_on=job_code_col, 
                                        suffixes=('_filtered', '_ref')
                                    )

                                    # Reset index of the result DataFrame
                                    result_dfpurifiers.reset_index(drop=True, inplace=True)

                                    # Check for title column with different possible names
                                    title_col = None
                                    for possible_title in ['Title', 'J3 Job Title', 'Task Description', 'Job Title']:
                                        if possible_title in result_dfpurifiers.columns:
                                            title_col = possible_title
                                            print(f"Found title column: {title_col}")
                                            break

                                    # Create pivot table if we have the necessary columns
                                    if title_col is not None and 'Machinery Location' in result_dfpurifiers.columns:
                                        pivot_table_resultpurifierJobs = result_dfpurifiers.pivot_table(
                                            index=title_col, 
                                            columns='Machinery Location', 
                                            values='Job Codecopy', 
                                            aggfunc='count'
                                            ).fillna(0).astype(int)  # Fill NaNs with 0s and convert to int

                                        # Apply color formatting only to numeric part
                                        # numeric_cols = pivot_table_resultpurifierJobs.select_dtypes(include=[np.number]).columns
                                        styled_pivotpurifier = pivot_table_resultpurifierJobs.style.applymap(color_binary_cells)

                                        # Display styled dataframe
                                        st.dataframe(
                                            styled_pivotpurifier,
                                            use_container_width=True,
                                            column_config={
                                                pivot_table_resultpurifierJobs.index.name or title_col: st.column_config.TextColumn(
                                                    label=title_col,
                                                    width="large"
                                                )
                                            }
                                        )

                                        # Provide download option
                                        csv_pivot = pivot_table_resultpurifierJobs.to_csv()
                                        st.download_button(
                                            label="Download Reference Jobs Pivot Table CSV",
                                            data=csv_pivot,
                                            file_name="purifier_ref_jobs_pivot.csv",
                                            mime="text/csv"
                                        )
                                else:
                                    st.info("Required columns for pivot table not found in merged data.")
                            except Exception as e:
                                st.error(f"Error creating purifier reference pivot table: {str(e)}")
                        else:
                            st.info("No reference job data available for purifiers.")

                    # Display Missing Jobs for Purifiers
                    st.subheader("Missing Jobs for Purifiers")
                    if not missing_jobs_purifier.empty:
                        # Use the HTML rendering approach to avoid JavaScript errors
                        st.write(missing_jobs_purifier.to_html(index=False), unsafe_allow_html=True)

                        # Provide download option
                        csv = missing_jobs_purifier.to_csv(index=False)
                        st.download_button(
                            label="Download Missing Jobs CSV",
                            data=csv,
                            file_name="purifier_missing_jobs.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No missing jobs analysis available for purifiers.")
                else:
                    st.warning("Reference sheet is required for purifier reference analysis. Please upload a reference sheet.")
            except Exception as e:
                st.error(f"Error displaying Reference Analysis: {str(e)}")

        elif st.session_state.current_tab == 4:
            st.header("Ballast Water Treatment System (BWTS) Analysis")

            # Initialize BWTSProcessor
            bwts_processor = BWTSProcessor()

            # BWTS Running Hours section removed as requested

            try:
                # Task Count Analysis
                st.subheader("Task Count Analysis for BWTS")
                task_count = bwts_processor.create_task_count_table(data)
                if not task_count.empty:
                    # Use HTML rendering approach to avoid JS errors
                    st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                    # Provide download option
                    csv = task_count.to_csv(index=False)
                    st.download_button(
                        label="Download Task Count CSV",
                        data=csv,
                        file_name="bwts_task_count.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No task count data available for BWTS.")
            except Exception as e:
                st.error(f"Error displaying Task Count Analysis: {str(e)}")

            try:
                # Reference Job Analysis
                if ref_sheet is not None:
                    # Get the selected BWTS model sheet name from session state
                    selected_bwts_model = st.session_state.bwts_model
                    selected_sheet_name = bwts_models[selected_bwts_model]

                    st.info(f"Using Reference Sheet: {selected_sheet_name} for {selected_bwts_model}")

                    # Get missing jobs data using the selected model's sheet name
                    missing_jobs_bwts = bwts_processor.process_reference_data(data, ref_sheet, preferred_sheet=selected_sheet_name)

                    # Generate result_dfbwts data
                    # Create a copy of data to avoid modifications to original
                    data_copy = data.copy()

                    # Filter for BWTS jobs with more flexible patterns
                    bwts_patterns = ['Ballast Water Treatment Plant', 'BWTS', 'Ballast Treatment']
                    mask = data_copy['Machinery Location'].str.contains('|'.join(bwts_patterns), case=False, na=False)
                    filtered_dfbwtsjobs = data_copy[mask].copy()

                    # Read the reference sheet
                    ref_sheet_names = pd.ExcelFile(ref_sheet).sheet_names

                    # First try to use the selected model sheet
                    bwts_sheet = selected_sheet_name if selected_sheet_name in ref_sheet_names else None

                    # If the selected sheet doesn't exist, fall back to default behavior
                    if bwts_sheet is None:
                        # Look for 'BWTS' sheet
                        bwts_sheet = 'BWTS' if 'BWTS' in ref_sheet_names else None

                        # If not found, try to find any sheet with 'BWTS', 'Ballast' or 'Water' in the name
                        if bwts_sheet is None:
                            for sheet in ref_sheet_names:
                                if 'BWTS' in sheet or 'Ballast' in sheet or 'Water' in sheet:
                                    bwts_sheet = sheet
                                    break

                        # If still no BWTS-specific sheet, use the first sheet
                        if bwts_sheet is None:
                            bwts_sheet = ref_sheet_names[0]
                            print(f"Selected sheet '{selected_sheet_name}' not found. No BWTS sheet found in app.py, using the first sheet: {bwts_sheet}")
                        else:
                            print(f"Selected sheet '{selected_sheet_name}' not found. Using reference sheet in app.py: {bwts_sheet}")
                    else:
                        print(f"Using selected model sheet in app.py: {bwts_sheet}")

                    # Read the reference sheet
                    dfbwts = pd.read_excel(ref_sheet, sheet_name=bwts_sheet)

                    # Create Job Codecopy column
                    if 'Job Code' in filtered_dfbwtsjobs.columns:
                        filtered_dfbwtsjobs['Job Codecopy'] = filtered_dfbwtsjobs['Job Code'].astype(str)

                        # Display Reference Jobs for BWTS
                        st.subheader("Reference Jobs for BWTS")
                        if not filtered_dfbwtsjobs.empty and not dfbwts.empty:
                            try:
                                # Display the columns in the reference data (for debugging)
                                print(f"Columns in reference data: {dfbwts.columns.tolist()}")

                                # Check which column to use for job code in reference data
                                job_code_col = None
                                for possible_col in ['UI Job Code', 'Job Code', 'JobCode', 'Code']:
                                    if possible_col in dfbwts.columns:
                                        job_code_col = possible_col
                                        print(f"Found job code column: {job_code_col}")
                                        break

                                if job_code_col is not None:
                                    # Convert job codes to string for proper matching
                                    dfbwts[job_code_col] = dfbwts[job_code_col].astype(str)

                                    # Print sample values for debugging
                                    print(f"Sample Job Codecopy values: {filtered_dfbwtsjobs['Job Codecopy'].iloc[:5].tolist()}")
                                    print(f"Sample {job_code_col} values: {dfbwts[job_code_col].iloc[:5].tolist()}")

                                    # Merge filtered_dfbwtsjobs with dfbwts on the matching job codes
                                    result_dfbwts = filtered_dfbwtsjobs.merge(dfbwts, left_on='Job Codecopy', right_on=job_code_col, suffixes=('_filtered', '_ref'))

                                    # Reset index of the result DataFrame
                                    result_dfbwts.reset_index(drop=True, inplace=True)

                                    # Check for title column with different possible names
                                    title_col = None
                                    for possible_title in ['Title', 'J3 Job Title', 'Task Description', 'Job Title']:
                                        if possible_title in result_dfbwts.columns:
                                            title_col = possible_title
                                            print(f"Found title column: {title_col}")
                                            break

                                    # Create pivot table if we have the necessary columns
                                    if title_col is not None and 'Machinery Location' in result_dfbwts.columns:
                                        pivot_table_resultbwtsJobs = result_dfbwts.pivot_table(
                                            index=title_col, 
                                            columns='Machinery Location', 
                                            values='Job Codecopy', 
                                            aggfunc='count'
                                        )

                                        # Display the pivot table
                                        st.write(pivot_table_resultbwtsJobs.to_html(), unsafe_allow_html=True)

                                        # Provide download option
                                        csv_pivot = pivot_table_resultbwtsJobs.to_csv()
                                        st.download_button(
                                            label="Download Reference Jobs Pivot Table CSV",
                                            data=csv_pivot,
                                            file_name="bwts_ref_jobs_pivot.csv",
                                            mime="text/csv"
                                        )
                                    else:
                                        st.info("Required columns for pivot table not found in merged data.")
                                else:
                                    st.info("No job code column found in reference data.")
                            except Exception as e:
                                st.error(f"Error creating BWTS reference pivot table: {str(e)}")
                        else:
                            st.info("No reference job data available for BWTS.")

                    # Display Missing Jobs for BWTS
                    st.subheader("Missing Jobs for BWTS")
                    if not missing_jobs_bwts.empty:
                        # Use the HTML rendering approach to avoid JavaScript errors
                        st.write(missing_jobs_bwts.to_html(index=False), unsafe_allow_html=True)

                        # Provide download option
                        csv = missing_jobs_bwts.to_csv(index=False)
                        st.download_button(
                            label="Download Missing Jobs CSV",
                            data=csv,
                            file_name="bwts_missing_jobs.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No missing jobs analysis available for BWTS.")
                else:
                    st.warning("Reference sheet is required for BWTS reference analysis. Please upload a reference sheet.")
            except Exception as e:
                st.error(f"Error displaying Reference Analysis: {str(e)}")

        elif st.session_state.current_tab == 5:
            st.header("Hatch Analysis")

            # Initialize HatchProcessor
            hatch_processor = HatchProcessor()

            try:
                # Task Count Analysis
                st.subheader("Task Count Analysis for Hatches")
                task_count = hatch_processor.create_task_count_table(data)
                if not task_count.empty:
                    # Use HTML rendering approach to avoid JS errors
                    st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                    # Provide download option
                    csv = task_count.to_csv(index=False)
                    st.download_button(
                        label="Download Task Count CSV",
                        data=csv,
                        file_name="hatch_task_count.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No task count data available for hatches.")
            except Exception as e:
                st.error(f"Error displaying Task Count Analysis: {str(e)}")



            try:
                # Reference Job Analysis
                if ref_sheet is not None:
                    # Create reference pivot table for Hatch Covers
                    pivot_table_hatch = hatch_processor.create_reference_pivot_table(data, ref_sheet)

                    # Display Reference Jobs Pivot Table for Hatch Covers 
                    st.subheader("Reference Jobs for Hatch Covers")
                    if not pivot_table_hatch.empty:
                        pivot_table_hatch = pivot_table_hatch.fillna(0).astype(int)

                        # Apply conditional formatting
                        styled_pivothatch = pivot_table_hatch.style.applymap(
                            color_binary_cells
                        )

                        # Display styled dataframe with wider first column
                        st.subheader("Reference Jobs for Hatch Covers")
                        st.dataframe(
                            styled_pivothatch,
                            use_container_width=True,
                            column_config={
                                pivot_table_hatch.index.name or pivot_table_hatch.columns[0]: st.column_config.TextColumn(
                                    label=pivot_table_hatch.index.name or pivot_table_hatch.columns[0],
                                    width="large"
                                )
                            }
                        )
                        # Provide download option for pivot table
                        csv_pivot = pivot_table_hatch.to_csv()
                        st.download_button(
                            label="Download Reference Jobs Pivot Table CSV",
                            data=csv_pivot,
                            file_name="hatch_ref_jobs_pivot.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No reference jobs pivot table available for hatches.")

                    # Get missing jobs data
                    missing_jobs_hatch = hatch_processor.process_reference_data(data, ref_sheet)

                    # Display Missing Jobs for Hatch Covers
                    st.subheader("Missing Jobs for Hatch Covers")
                    if not missing_jobs_hatch.empty:
                        # Use the HTML rendering approach to avoid JavaScript errors
                        st.write(missing_jobs_hatch.to_html(index=False), unsafe_allow_html=True)

                        # Provide download option
                        csv = missing_jobs_hatch.to_csv(index=False)
                        st.download_button(
                            label="Download Missing Jobs CSV",
                            data=csv,
                            file_name="hatch_missing_jobs.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No missing jobs analysis available for hatches.")
                else:
                    st.warning("Reference sheet is required for hatch reference analysis. Please upload a reference sheet.")
            except Exception as e:
                st.error(f"Error displaying Reference Analysis: {str(e)}")

        elif st.session_state.current_tab == 6:
            st.header("Cargo Pumping System Analysis")

            # Initialize processor
            cargopump_processor = CargoPumpingProcessor()

            # First: process data with reference
            if ref_sheet is not None:
                matched_jobs = cargopump_processor.process_reference_data(data, ref_sheet)

                try:
                    # Task Count Analysis
                    st.subheader("Task Count Analysis for Cargo Pumping")
                    task_count = cargopump_processor.create_task_count_table()
                    if task_count is not None and hasattr(task_count, 'to_html'):
                        st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                        csv = cargopump_processor.result_dfcargopumping.to_csv(index=False)
                        st.download_button(
                            label="Download Task Count CSV",
                            data=csv,
                            file_name="cargo_pumping_task_count.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No task count data available.")
                except Exception as e:
                    st.error(f"Error displaying Task Count Analysis: {str(e)}")

                st.subheader("Reference Jobs for Cargo Pumping (Matching Jobs)")
                if matched_jobs is not None and not matched_jobs.empty:
                    required_columns = ['UI Job Code', 'J3 Job Title', 'Remarks', 'Applicability']
                    filtered_matched_jobs = matched_jobs[required_columns] if all(
                        col in matched_jobs.columns for col in required_columns
                    ) else matched_jobs

                    # Remove duplicates by UI Job Code
                    filtered_matched_jobs = filtered_matched_jobs.drop_duplicates(subset=['UI Job Code'])

                    st.write(filtered_matched_jobs.to_html(index=False), unsafe_allow_html=True)

                    csv = filtered_matched_jobs.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Jobs CSV",
                        data=csv,
                        file_name="cargo_pumping_matched_jobs.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No matching jobs found for Cargo Pumping.")

                st.subheader("Missing Jobs for Cargo Pumping")
                missing_jobs = cargopump_processor.missingjobscargopumpingresult
                if not missing_jobs.empty:
                    st.write(missing_jobs.to_html(index=False), unsafe_allow_html=True)

                    csv = missing_jobs.to_csv(index=False)
                    st.download_button(
                        label="Download Missing Jobs CSV",
                        data=csv,
                        file_name="cargo_pumping_missing_jobs.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No missing job codes for Cargo Pumping.")
            else:
                st.warning("Reference sheet is required for Cargo Pumping reference analysis. Please upload a reference sheet.")

        elif st.session_state.current_tab == 7:
            st.header("Inert Gas System Analysis")

            ig_processor = InertGasSystemProcessor()

            if ref_sheet is not None:
                matched_jobs = ig_processor.process_reference_data(data, ref_sheet)

                try:
                    st.subheader("Task Count Analysis for Inert Gas System")
                    task_count = ig_processor.create_task_count_table()
                    if task_count is not None and hasattr(task_count, 'to_html'):
                        st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                        csv = ig_processor.result_df_igsystem.to_csv(index=False)
                        st.download_button(
                            label="Download Task Count CSV",
                            data=csv,
                            file_name="igsystem_task_count.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No task count data available.")
                except Exception as e:
                    st.error(f"Error displaying Task Count Analysis: {str(e)}")

                st.subheader("Reference Jobs for Inert Gas System (Matching Jobs)")
                if matched_jobs is not None and not matched_jobs.empty:
                    required_columns = ['UI Job Code', 'J3 Job Title', 'Remarks', 'Applicability']
                    filtered_matched_jobs = matched_jobs[required_columns] if all(
                        col in matched_jobs.columns for col in required_columns
                    ) else matched_jobs

                    # Remove duplicates by UI Job Code
                    filtered_matched_jobs = filtered_matched_jobs.drop_duplicates(subset=['UI Job Code'])

                    st.write(filtered_matched_jobs.to_html(index=False), unsafe_allow_html=True)

                    csv = filtered_matched_jobs.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Jobs CSV",
                        data=csv,
                        file_name="igsystem_matched_jobs.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No matching jobs found for Inert Gas System.")

                st.subheader("Missing Jobs for Inert Gas System")
                missing_jobs = ig_processor.missing_jobs_igsystem
                if not missing_jobs.empty:
                    st.dataframe(missing_jobs)

                    csv = missing_jobs.to_csv(index=False)
                    st.download_button(
                        label="Download Missing Jobs CSV",
                        data=csv,
                        file_name="igsystem_missing_jobs.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No missing jobs from reference for Inert Gas System.")
            else:
                st.warning("Reference sheet is required for Inert Gas System reference analysis.")

        elif st.session_state.current_tab == 8:
            st.header("Cargo Handling System Analysis")

            chs_processor = CargoHandlingSystemProcessor()

            if ref_sheet is not None:
                matched_jobs = chs_processor.process_reference_data(data, ref_sheet)

                try:
                    st.subheader("Task Count Analysis for Cargo Handling System")
                    task_count = chs_processor.create_task_count_table()
                    if task_count is not None and hasattr(task_count, 'to_html'):
                        st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                        csv = chs_processor.result_df_cargohandling.to_csv(index=False)
                        st.download_button(
                            label="Download Task Count CSV",
                            data=csv,
                            file_name="cargohandling_task_count.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No task count data available.")
                except Exception as e:
                    st.error(f"Error displaying Task Count Analysis: {str(e)}")

                st.subheader("Reference Jobs for Cargo Handling System (Matching Jobs)")
                if matched_jobs is not None and not matched_jobs.empty:
                    required_columns = ['UI Job Code', 'J3 Job Title', 'Remarks', 'Applicability']
                    filtered_matched_jobs = matched_jobs[required_columns] if all(
                        col in matched_jobs.columns for col in required_columns
                    ) else matched_jobs

                    # Remove duplicates by UI Job Code
                    filtered_matched_jobs = filtered_matched_jobs.drop_duplicates(subset=['UI Job Code'])

                    st.write(filtered_matched_jobs.to_html(index=False), unsafe_allow_html=True)

                    csv = filtered_matched_jobs.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Jobs CSV",
                        data=csv,
                        file_name="cargohandling_matched_jobs.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No matching jobs found for Cargo Handling System.")

                st.subheader("Missing Jobs for Cargo Handling System")
                missing_jobs = chs_processor.missing_jobs_cargohandling
                if not missing_jobs.empty:
                    st.dataframe(missing_jobs)

                    csv = missing_jobs.to_csv(index=False)
                    st.download_button(
                        label="Download Missing Jobs CSV",
                        data=csv,
                        file_name="cargohandling_missing_jobs.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No missing jobs from reference for Cargo Handling System.")
            else:
                st.warning("Reference sheet is required for Cargo Handling System reference analysis.")


        elif st.session_state.current_tab == 9:
            st.header("Cargo Venting System Analysis")

            cvs_processor = CargoVentingSystemProcessor()

            if ref_sheet is not None:
                matched_jobs = cvs_processor.process_reference_data(data, ref_sheet)

                try:
                    st.subheader("Task Count Analysis for Cargo Venting System")
                    task_count = cvs_processor.create_task_count_table()
                    if task_count is not None and hasattr(task_count, 'to_html'):
                        st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                        csv = cvs_processor.result_df_cargovent.to_csv(index=False)
                        st.download_button(
                            label="Download Task Count CSV",
                            data=csv,
                            file_name="cargoventing_task_count.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No task count data available.")
                except Exception as e:
                    st.error(f"Error displaying Task Count Analysis: {str(e)}")

                st.subheader("Reference Jobs for Cargo Venting System (Matching Jobs)")
                if matched_jobs is not None and not matched_jobs.empty:
                    required_columns = ['UI Job Code', 'J3 Job Title', 'Remarks', 'Applicability']
                    filtered_matched_jobs = matched_jobs[required_columns] if all(
                        col in matched_jobs.columns for col in required_columns
                    ) else matched_jobs

                    # Remove duplicates by UI Job Code
                    filtered_matched_jobs = filtered_matched_jobs.drop_duplicates(subset=['UI Job Code'])

                    st.write(filtered_matched_jobs.to_html(index=False), unsafe_allow_html=True)

                    csv = filtered_matched_jobs.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Jobs CSV",
                        data=csv,
                        file_name="cargoventing_matched_jobs.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No matching jobs found for Cargo Venting System.")

                st.subheader("Missing Jobs for Cargo Venting System")
                missing_jobs = cvs_processor.missing_jobs_cargovent
                if not missing_jobs.empty:
                    st.dataframe(missing_jobs)

                    csv = missing_jobs.to_csv(index=False)
                    st.download_button(
                        label="Download Missing Jobs CSV",
                        data=csv,
                        file_name="cargoventing_missing_jobs.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No missing jobs from reference for Cargo Venting System.")
            else:
                st.warning("Reference sheet is required for Cargo Venting System reference analysis.")


        elif st.session_state.current_tab == 10:
                st.header("LSA/FFA System Analysis")

                lsaffa_processor = LSAFFAProcessor()

                if ref_sheet is not None:
                    matched_jobs = lsaffa_processor.process_reference_data(data, ref_sheet)

                    try:
                        st.subheader("Task Count Analysis for LSA/FFA")
                        task_count = lsaffa_processor.create_task_count_table()
                        if task_count is not None and hasattr(task_count, 'to_html'):
                            st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                            csv = lsaffa_processor.result_df_lsaffa.to_csv(index=False)
                            st.download_button(
                                label="Download Task Count CSV",
                                data=csv,
                                file_name="lsaffa_task_count.csv",
                                mime="text/csv"
                            )
                        else:
                            st.info("No task count data available.")
                    except Exception as e:
                        st.error(f"Error displaying Task Count Analysis: {str(e)}")

                    st.subheader("Reference Jobs for LSA/FFA (Matching Jobs)")
                    if not matched_jobs.empty:
                        display_cols = ['Machinery', 'UI Job Code', 'J3 Job Title', 'Remarks', 'Applicability']
                        filtered_matched_jobs = matched_jobs[[col for col in display_cols if col in matched_jobs.columns]]

                        st.write(filtered_matched_jobs.to_html(index=False), unsafe_allow_html=True)
                        csv = matched_jobs.to_csv(index=False)
                        st.download_button(
                            label="Download Matched Jobs CSV",
                            data=csv,
                            file_name="lsaffa_matched_jobs.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No matching jobs found for LSA/FFA.")

                    st.subheader("Missing Jobs for LSA/FFA")
                    missing_jobs = lsaffa_processor.missing_jobs_lsaffa
                    if not missing_jobs.empty:
                        st.dataframe(missing_jobs)

                        csv = missing_jobs.to_csv(index=False)
                        st.download_button(
                            label="Download Missing Jobs CSV",
                            data=csv,
                            file_name="lsaffa_missing_jobs.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No missing jobs from reference for LSA/FFA.")
                else:
                    st.warning("Reference sheet is required for LSA/FFA analysis.")

        elif st.session_state.current_tab == 11:
                st.header("Fire Fighting System Analysis")

                ffasys_processor = FFASystemProcessor()

                if ref_sheet is not None:
                    matched_jobs = ffasys_processor.process_reference_data(data, ref_sheet)

                    try:
                        st.subheader("Task Count Analysis for Fire Fighting System")
                        task_count = ffasys_processor.create_task_count_table()
                        if task_count is not None and hasattr(task_count, 'to_html'):
                            st.write(task_count.to_html(index=False), unsafe_allow_html=True)

                            csv = ffasys_processor.result_df_ffasys.to_csv(index=False)
                            st.download_button(
                                label="Download Task Count CSV",
                                data=csv,
                                file_name="ffasys_task_count.csv",
                                mime="text/csv"
                            )
                        else:
                            st.info("No task count data available.")
                    except Exception as e:
                        st.error(f"Error displaying Task Count Analysis: {str(e)}")

                    st.subheader("Reference Jobs for Fire Fighting System (Matching Jobs)")
                    if not matched_jobs.empty:
                        display_cols = ['Machinery', 'UI Job Code', 'J3 Job Title', 'Remarks', 'Applicability']
                        filtered_matched_jobs = matched_jobs[[col for col in display_cols if col in matched_jobs.columns]]

                        st.write(filtered_matched_jobs.to_html(index=False), unsafe_allow_html=True)

                        csv = matched_jobs.to_csv(index=False)
                        st.download_button(
                            label="Download Matched Jobs CSV",
                            data=csv,
                            file_name="ffasys_matched_jobs.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No matching jobs found for Fire Fighting System.")

                    st.subheader("Missing Jobs for Fire Fighting System")
                    missing_jobs = ffasys_processor.missing_jobs_ffasys
                    if not missing_jobs.empty:
                        st.dataframe(missing_jobs)

                        csv = missing_jobs.to_csv(index=False)
                        st.download_button(
                            label="Download Missing Jobs CSV",
                            data=csv,
                            file_name="ffasys_missing_jobs.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No missing jobs from reference for Fire Fighting System.")
                else:
                    st.warning("Reference sheet is required for Fire Fighting System analysis.")


        elif st.session_state.current_tab == 12:
                st.header("Pump System Analysis")

                pump_processor = PumpSystemProcessor()

                if ref_sheet is not None:
                    try:
                        dfpump = pd.read_excel(ref_sheet, sheet_name='Pumps')
                        pump_output = pump_processor.process_pump_data(data, dfpump)

                        # 🔹 Display Pump Count by Location
                        st.subheader("Pump Location Task Count")
                        if pump_processor.styled_pivot_table_pump is not None:
                            st.markdown(pump_processor.styled_pivot_table_pump.to_html(), unsafe_allow_html=True)
                        else:
                            st.dataframe(pump_processor.pivot_table_pump)

                        # 🔹 Display Mapped Job Code Summary Table
                        st.subheader("Mapped Job Code Summary Table")
                        if pump_processor.styled_pivot_table_resultpumpJobs is not None:
                            st.markdown(pump_processor.styled_pivot_table_resultpumpJobs.to_html(), unsafe_allow_html=True)
                        elif not pump_processor.pivot_table_resultpumpJobs.empty:
                            st.dataframe(pump_processor.pivot_table_resultpumpJobs)
                        else:
                            st.info("No mapped pump jobs found to display.")

                        # 🔹 Download Buttons
                        csv1 = pump_processor.filtered_dfpump.to_csv(index=False)
                        st.download_button(
                            label="Download Mapped Pump Jobs CSV",
                            data=csv1,
                            file_name="mapped_pump_jobs.csv",
                            mime="text/csv"
                        )

                        csv2 = pump_processor.pivot_table_resultpumpJobs.to_csv()
                        st.download_button(
                            label="Download Job Code Summary CSV",
                            data=csv2,
                            file_name="pump_summary_table.csv",
                            mime="text/csv"
                        )

                    except Exception as e:
                        st.error(f"Error processing Pump data: {str(e)}")
                else:
                    st.warning("Reference sheet is required for Pump analysis.")


        elif st.session_state.current_tab == 13:
            st.header("Compressor System Analysis")

            compressor_processor = CompressorSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfCompressor = pd.read_excel(ref_sheet, sheet_name='Compressor')
                    compressor_processor.process_compressor_data(data, dfCompressor)

                    st.subheader("Matched Compressor Job Code Summary Table")
                    if compressor_processor.pivot_table_resultCompressorJobs is not None and not compressor_processor.pivot_table_resultCompressorJobs.empty:
                        styled_compressor = compressor_processor.pivot_table_resultCompressorJobs.style.applymap(color_binary_cells)

                        st.dataframe(
                            styled_compressor,
                            use_container_width=True,
                            column_config={
                                compressor_processor.pivot_table_resultCompressorJobs.index.name or compressor_processor.pivot_table_resultCompressorJobs.columns[0]: st.column_config.TextColumn(
                                    label=compressor_processor.pivot_table_resultCompressorJobs.index.name or compressor_processor.pivot_table_resultCompressorJobs.columns[0],
                                    width="large"
                                )
                            }
                        )
                    else:
                        st.info("No mapped compressor jobs found to display.")

                    st.subheader("Missing Job Codes from Compressor Reference")
                    if not compressor_processor.missingjobsCompressorresult.empty:
                        st.dataframe(compressor_processor.missingjobsCompressorresult)
                    else:
                        st.success("All job codes are matched in the Compressor reference sheet.")

                    # Downloads
                    csv1 = compressor_processor.result_dfCompressor.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Compressor Jobs CSV",
                        data=csv1,
                        file_name="matched_compressor_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = compressor_processor.pivot_table_resultCompressorJobs.to_csv()
                    st.download_button(
                        label="Download Compressor Summary Table CSV",
                        data=csv2,
                        file_name="compressor_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Compressor data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Compressor analysis.")

        elif st.session_state.current_tab == 14:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Ladder System Analysis")

            ladder_processor = LadderSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfLadder = pd.read_excel(ref_sheet, sheet_name='Ladders')
                    ladder_processor.process_ladder_data(data, dfLadder)

                    st.subheader("Matched Ladder Job Code Summary Table")
                    if ladder_processor.pivot_table_resultLadderJobs is not None and not ladder_processor.pivot_table_resultLadderJobs.empty:
                        styled_ladder = ladder_processor.pivot_table_resultLadderJobs.style.applymap(color_binary_cells)

                        st.dataframe(
                            styled_ladder,
                            use_container_width=True,
                            column_config={
                                ladder_processor.pivot_table_resultLadderJobs.index.name or ladder_processor.pivot_table_resultLadderJobs.columns[0]: st.column_config.TextColumn(
                                    label=ladder_processor.pivot_table_resultLadderJobs.index.name or ladder_processor.pivot_table_resultLadderJobs.columns[0],
                                    width="large"
                                )
                            }
                        )
                    else:
                        st.info("No mapped ladder jobs found to display.")

                    st.subheader("Missing Job Codes from Ladder Reference")
                    if not ladder_processor.missingjobsLadderresult.empty:
                        st.dataframe(ladder_processor.missingjobsLadderresult)
                    else:
                        st.success("All job codes are matched in the Ladder reference sheet.")

                    # Downloads
                    csv1 = ladder_processor.result_dfLadder.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Ladder Jobs CSV",
                        data=csv1,
                        file_name="matched_ladder_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = ladder_processor.pivot_table_resultLadderJobs.to_csv()
                    st.download_button(
                        label="Download Ladder Summary Table CSV",
                        data=csv2,
                        file_name="ladder_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Ladder data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Ladder analysis.")


        elif st.session_state.current_tab == 15:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Boat System Analysis")

            boat_processor = BoatSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfBoats = pd.read_excel(ref_sheet, sheet_name='Boats')
                    boat_processor.process_boat_data(data, dfBoats)

                    st.subheader("Matched Boat Job Code Summary Table")
                    if boat_processor.pivot_table_resultBoatJobs is not None and not boat_processor.pivot_table_resultBoatJobs.empty:
                        styled_boat = boat_processor.pivot_table_resultBoatJobs.style.applymap(color_binary_cells)

                        st.dataframe(
                            styled_boat,
                            use_container_width=True,
                            column_config={
                                boat_processor.pivot_table_resultBoatJobs.index.name or boat_processor.pivot_table_resultBoatJobs.columns[0]: st.column_config.TextColumn(
                                    label=boat_processor.pivot_table_resultBoatJobs.index.name or boat_processor.pivot_table_resultBoatJobs.columns[0],
                                    width="large"
                                )
                            }
                        )
                    else:
                        st.info("No mapped boat jobs found to display.")

                    st.subheader("Missing Job Codes from Boat Reference")
                    if not boat_processor.missingjobsBoatsresult.empty:
                        st.dataframe(boat_processor.missingjobsBoatsresult)
                    else:
                        st.success("All job codes are matched in the Boat reference sheet.")

                    # Downloads
                    csv1 = boat_processor.result_dfBoat.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Boat Jobs CSV",
                        data=csv1,
                        file_name="matched_boat_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = boat_processor.pivot_table_resultBoatJobs.to_csv()
                    st.download_button(
                        label="Download Boat Summary Table CSV",
                        data=csv2,
                        file_name="boat_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Boat data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Boat analysis.")


        elif st.session_state.current_tab == 16:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Mooring System Analysis")

            mooring_processor = MooringSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfMooring = pd.read_excel(ref_sheet, sheet_name='Mooring')
                    mooring_processor.process_mooring_data(data, dfMooring)

                    st.subheader("Matched Mooring Job Code Summary Table")
                    if mooring_processor.pivot_table_resultMooringJobs is not None and not mooring_processor.pivot_table_resultMooringJobs.empty:
                        styled_mooring = mooring_processor.pivot_table_resultMooringJobs.style.applymap(color_binary_cells)

                        st.dataframe(
                            styled_mooring,
                            use_container_width=True,
                            column_config={
                                mooring_processor.pivot_table_resultMooringJobs.index.name or mooring_processor.pivot_table_resultMooringJobs.columns[0]: st.column_config.TextColumn(
                                    label=mooring_processor.pivot_table_resultMooringJobs.index.name or mooring_processor.pivot_table_resultMooringJobs.columns[0],
                                    width="large"
                                )
                            }
                        )
                    else:
                        st.info("No mapped mooring jobs found to display.")

                    st.subheader("Missing Job Codes from Mooring Reference")
                    if not mooring_processor.missingjobsMooringresult.empty:
                        st.dataframe(mooring_processor.missingjobsMooringresult)
                    else:
                        st.success("All job codes are matched in the Mooring reference sheet.")

                    # Downloads
                    csv1 = mooring_processor.result_dfMooring.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Mooring Jobs CSV",
                        data=csv1,
                        file_name="matched_mooring_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = mooring_processor.pivot_table_resultMooringJobs.to_csv()
                    st.download_button(
                        label="Download Mooring Summary Table CSV",
                        data=csv2,
                        file_name="mooring_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Mooring data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Mooring analysis.")



        elif st.session_state.current_tab == 17:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Steering System Analysis")

            steering_processor = SteeringSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfSteering = pd.read_excel(ref_sheet, sheet_name='Steering')
                    steering_processor.process_steering_data(data, dfSteering)

                    st.subheader("Matched Steering Job Code Summary Table")
                    if steering_processor.pivot_table_resultSteeringJobs is not None and not steering_processor.pivot_table_resultSteeringJobs.empty:
                        styled_steering = steering_processor.pivot_table_resultSteeringJobs.style.applymap(color_binary_cells)

                        st.dataframe(
                            styled_steering,
                            use_container_width=True,
                            column_config={
                                steering_processor.pivot_table_resultSteeringJobs.index.name or steering_processor.pivot_table_resultSteeringJobs.columns[0]: st.column_config.TextColumn(
                                    label=steering_processor.pivot_table_resultSteeringJobs.index.name or steering_processor.pivot_table_resultSteeringJobs.columns[0],
                                    width="large"
                                )
                            }
                        )
                    else:
                        st.info("No mapped steering jobs found to display.")

                    st.subheader("Missing Job Codes from Steering Reference")
                    if not steering_processor.missingjobsSteeringresult.empty:
                        st.dataframe(steering_processor.missingjobsSteeringresult)
                    else:
                        st.success("All job codes are matched in the Steering reference sheet.")

                    # Downloads
                    csv1 = steering_processor.result_dfSteering.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Steering Jobs CSV",
                        data=csv1,
                        file_name="matched_steering_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = steering_processor.pivot_table_resultSteeringJobs.to_csv()
                    st.download_button(
                        label="Download Steering Summary Table CSV",
                        data=csv2,
                        file_name="steering_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Steering data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Steering analysis.")

        elif st.session_state.current_tab == 18:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Incinerator System Analysis")

            incin_processor = IncineratorSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfIncin = pd.read_excel(ref_sheet, sheet_name='Incin')
                    incin_processor.process_incin_data(data, dfIncin)

                    st.subheader("Matched Incinerator Job Code Summary Table")
                    if incin_processor.pivot_table_resultIncinJobs is not None and not incin_processor.pivot_table_resultIncinJobs.empty:
                        styled_incin = incin_processor.pivot_table_resultIncinJobs.style.applymap(color_binary_cells)

                        st.dataframe(
                            styled_incin,
                            use_container_width=True,
                            column_config={
                                incin_processor.pivot_table_resultIncinJobs.index.name or incin_processor.pivot_table_resultIncinJobs.columns[0]: st.column_config.TextColumn(
                                    label=incin_processor.pivot_table_resultIncinJobs.index.name or incin_processor.pivot_table_resultIncinJobs.columns[0],
                                    width="large"
                                )
                            }
                        )
                    else:
                        st.info("No mapped incinerator jobs found to display.")

                    st.subheader("Missing Job Codes from Incinerator Reference")
                    if not incin_processor.missingjobsIncinresult.empty:
                        st.dataframe(incin_processor.missingjobsIncinresult)
                    else:
                        st.success("All job codes are matched in the Incinerator reference sheet.")

                    # Downloads
                    csv1 = incin_processor.result_dfIncin.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Incinerator Jobs CSV",
                        data=csv1,
                        file_name="matched_incinerator_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = incin_processor.pivot_table_resultIncinJobs.to_csv()
                    st.download_button(
                        label="Download Incinerator Summary Table CSV",
                        data=csv2,
                        file_name="incinerator_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Incinerator data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Incinerator analysis.")

        elif st.session_state.current_tab == 19:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("STP System Analysis")

            stp_processor = STPSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfSTP = pd.read_excel(ref_sheet, sheet_name='STP')
                    stp_processor.process_stp_data(data, dfSTP)

                    st.subheader("Matched STP Job Code Summary Table")

                    if stp_processor.pivot_table_resultSTPJobs is not None and not stp_processor.pivot_table_resultSTPJobs.empty:
                        styled_stp = stp_processor.pivot_table_resultSTPJobs.style.applymap(color_binary_cells)

                        st.dataframe(
                            styled_stp,
                            use_container_width=True,
                            column_config={
                                stp_processor.pivot_table_resultSTPJobs.index.name or stp_processor.pivot_table_resultSTPJobs.columns[0]: st.column_config.TextColumn(
                                    label=stp_processor.pivot_table_resultSTPJobs.index.name or stp_processor.pivot_table_resultSTPJobs.columns[0],
                                    width="large"
                                )
                            }
                        )
                    else:
                        st.info("No mapped STP jobs found to display.")

                    st.subheader("Missing Job Codes from STP Reference")
                    if not stp_processor.missingjobsSTPresult.empty:
                        st.dataframe(stp_processor.missingjobsSTPresult)
                    else:
                        st.success("All job codes are matched in the STP reference sheet.")

                    # Downloads
                    csv1 = stp_processor.result_dfSTP.to_csv(index=False)
                    st.download_button(
                        label="Download Matched STP Jobs CSV",
                        data=csv1,
                        file_name="matched_stp_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = stp_processor.pivot_table_resultSTPJobs.to_csv()
                    st.download_button(
                        label="Download STP Summary Table CSV",
                        data=csv2,
                        file_name="stp_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing STP data: {str(e)}")
            else:
                st.warning("Reference sheet is required for STP analysis.")



        elif st.session_state.current_tab == 20:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("OWS System Analysis")

            ows_processor = OWSSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfOWS = pd.read_excel(ref_sheet, sheet_name='OWS')
                    ows_processor.process_ows_data(data, dfOWS)

                    st.subheader("Matched OWS Job Code Summary Table")
                    if ows_processor.pivot_table_resultOWSJobs is not None and not ows_processor.pivot_table_resultOWSJobs.empty:
                        styled_ows = ows_processor.pivot_table_resultOWSJobs.style.applymap(color_binary_cells)

                        st.dataframe(
                            styled_ows,
                            use_container_width=True,
                            column_config={
                                ows_processor.pivot_table_resultOWSJobs.index.name or ows_processor.pivot_table_resultOWSJobs.columns[0]: st.column_config.TextColumn(
                                    label=ows_processor.pivot_table_resultOWSJobs.index.name or ows_processor.pivot_table_resultOWSJobs.columns[0],
                                    width="large"
                                )
                            }
                        )
                    else:
                        st.info("No mapped OWS jobs found to display.")

                    st.subheader("Missing Job Codes from OWS Reference")
                    if not ows_processor.missingjobsOWSresult.empty:
                        st.dataframe(ows_processor.missingjobsOWSresult)
                    else:
                        st.success("All job codes are matched in the OWS reference sheet.")

                    # Downloads
                    csv1 = ows_processor.result_dfOWS.to_csv(index=False)
                    st.download_button(
                        label="Download Matched OWS Jobs CSV",
                        data=csv1,
                        file_name="matched_ows_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = ows_processor.pivot_table_resultOWSJobs.to_csv()
                    st.download_button(
                        label="Download OWS Summary Table CSV",
                        data=csv2,
                        file_name="ows_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing OWS data: {str(e)}")
            else:
                st.warning("Reference sheet is required for OWS analysis.")

        elif st.session_state.current_tab == 21:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Power Distribution System Analysis")

            powerdist_processor = PowerDistSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfpowerdist = pd.read_excel(ref_sheet, sheet_name='Powerdist')
                    powerdist_processor.process_powerdist_data(data, dfpowerdist)

                    st.subheader("Matched Power Distribution Job Code Summary Table")
                    pivot_df = powerdist_processor.pivot_table_resultpowerdistJobs  # ✅ Correct casing

                    if pivot_df is not None and not pivot_df.empty:
                        styled_powerdist = pivot_df.style.applymap(color_binary_cells)

                        first_col = pivot_df.index.name or (pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1")

                        st.dataframe(
                            styled_powerdist,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped Power Distribution jobs found to display.")

                    st.subheader("Missing Job Codes from Power Distribution Reference")
                    if not powerdist_processor.missingjobspowerdistresult.empty:
                        st.dataframe(powerdist_processor.missingjobspowerdistresult)
                    else:
                        st.success("All job codes are matched in the Power Distribution reference sheet.")

                    # Downloads
                    csv1 = powerdist_processor.result_dfpowerdist.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Power Distribution Jobs CSV",
                        data=csv1,
                        file_name="matched_powerdist_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = powerdist_processor.pivot_table_resultpowerdistJobs.to_csv()
                    st.download_button(
                        label="Download Power Distribution Summary Table CSV",
                        data=csv2,
                        file_name="powerdist_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Power Distribution data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Power Distribution analysis.")

        elif st.session_state.current_tab == 22:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Crane System Analysis")

            crane_processor = CraneSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfcrane = pd.read_excel(ref_sheet, sheet_name='Crane')
                    crane_processor.process_crane_data(data, dfcrane)

                    st.subheader("Matched Crane Job Code Summary Table")
                    pivot_df = crane_processor.pivot_table_resultcraneJobs

                    if pivot_df is not None and not pivot_df.empty:
                        styled_crane = pivot_df.style.applymap(color_binary_cells)

                        first_col = pivot_df.index.name or (pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1")

                        st.dataframe(
                            styled_crane,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped crane jobs found to display.")

                    st.subheader("Missing Job Codes from Crane Reference")
                    if not crane_processor.missingjobscraneresult.empty:
                        st.dataframe(crane_processor.missingjobscraneresult)
                    else:
                        st.success("All job codes are matched in the Crane reference sheet.")

                    # Downloads
                    csv1 = crane_processor.result_dfcrane.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Crane Jobs CSV",
                        data=csv1,
                        file_name="matched_crane_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = crane_processor.pivot_table_resultcraneJobs.to_csv()
                    st.download_button(
                        label="Download Crane Summary Table CSV",
                        data=csv2,
                        file_name="crane_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Crane data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Crane analysis.")



        elif st.session_state.current_tab == 23:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Emergency Generator System Analysis")

            emg_processor = EmergencyGenSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfEmg = pd.read_excel(ref_sheet, sheet_name='Emg')
                    emg_processor.process_emg_data(data, dfEmg)

                    st.subheader("Matched Emergency Generator Job Code Summary Table")
                    pivot_df = emg_processor.pivot_table_resultEmgJobs  # ✅ Correct attribute name

                    if pivot_df is not None and not pivot_df.empty:
                        styled_emg = pivot_df.style.applymap(color_binary_cells)

                        first_col = pivot_df.index.name or (pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1")

                        st.dataframe(
                            styled_emg,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped emergency generator jobs found to display.")

                    st.subheader("Missing Job Codes from Emergency Generator Reference")
                    if not emg_processor.missingjobsEmgresult.empty:
                        st.dataframe(emg_processor.missingjobsEmgresult)
                    else:
                        st.success("All job codes are matched in the Emergency Generator reference sheet.")

                    # Downloads
                    csv1 = emg_processor.result_dfEmg.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Emergency Generator Jobs CSV",
                        data=csv1,
                        file_name="matched_emg_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = emg_processor.pivot_table_resultEmgJobs.to_csv()
                    st.download_button(
                        label="Download Emergency Generator Summary Table CSV",
                        data=csv2,
                        file_name="emg_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Emergency Generator data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Emergency Generator analysis.")


        elif st.session_state.current_tab == 24:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Bridge System Analysis")

            bridge_processor = BridgeSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfbridge = pd.read_excel(ref_sheet, sheet_name='Bridge')
                    bridge_processor.process_bridge_data(data, dfbridge)

                    st.subheader("Matched Bridge Job Code Summary Table")
                    pivot_df = bridge_processor.pivot_table_resultbridgeJobs  # ✅ correct attribute name

                    if pivot_df is not None and not pivot_df.empty:
                        styled_bridge = pivot_df.style.applymap(color_binary_cells)

                        first_col = pivot_df.index.name or (pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1")

                        st.dataframe(
                            styled_bridge,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped bridge jobs found to display.")

                    st.subheader("Missing Job Codes from Bridge Reference")
                    if not bridge_processor.missingjobsbridgeresult.empty:
                        st.dataframe(bridge_processor.missingjobsbridgeresult)
                    else:
                        st.success("All job codes are matched in the Bridge reference sheet.")

                    # Downloads
                    csv1 = bridge_processor.result_dfbridge.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Bridge Jobs CSV",
                        data=csv1,
                        file_name="matched_bridge_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = bridge_processor.pivot_table_resultbridgeJobs.to_csv()
                    st.download_button(
                        label="Download Bridge Summary Table CSV",
                        data=csv2,
                        file_name="bridge_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Bridge data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Bridge analysis.")


        elif st.session_state.current_tab == 25:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Reefer & AC System Analysis")

            refac_processor = RefacSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfrefac = pd.read_excel(ref_sheet, sheet_name='Refac')
                    refac_processor.process_refac_data(data, dfrefac)

                    st.subheader("Matched Reefer & AC Job Code Summary Table")
                    pivot_df = refac_processor.pivot_table_resultrefacJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_refac = pivot_df.style.applymap(color_binary_cells)

                        first_col = pivot_df.index.name or (pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1")

                        st.dataframe(
                            styled_refac,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped Reefer & AC jobs found to display.")

                    st.subheader("Missing Job Codes from Reefer & AC Reference")
                    if not refac_processor.missingjobsrefacresult.empty:
                        st.dataframe(refac_processor.missingjobsrefacresult)
                    else:
                        st.success("All job codes are matched in the Reefer & AC reference sheet.")

                    # Downloads
                    csv1 = refac_processor.result_dfrefac.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Reefer & AC Jobs CSV",
                        data=csv1,
                        file_name="matched_refac_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = refac_processor.pivot_table_resultrefacJobs.to_csv()
                    st.download_button(
                        label="Download Reefer & AC Summary Table CSV",
                        data=csv2,
                        file_name="refac_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Refac data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Reefer & AC analysis.")

        elif st.session_state.current_tab == 26:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Fan System Analysis")

            fan_processor = FanSystemProcessor()

            if ref_sheet is not None:
                try:
                    dffan = pd.read_excel(ref_sheet, sheet_name='Fans')
                    fan_processor.process_fan_data(data, dffan)

                    st.subheader("Matched Fan Job Code Summary Table (By Title)")
                    pivot_df = fan_processor.pivot_table_resultfanJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_fan = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_fan,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No matched fan jobs found to display.")

                    st.subheader("Total Fan Jobs by Machinery and Sub Component")
                    if fan_processor.styled_pivot_table_fan is not None:
                        st.markdown(
                            fan_processor.styled_pivot_table_fan.to_html(),
                            unsafe_allow_html=True
                        )
                    else:
                        st.dataframe(fan_processor.pivot_table_fan)

                    # Downloads
                    csv1 = fan_processor.matching_jobsfan.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Fan Jobs CSV",
                        data=csv1,
                        file_name="matched_fan_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = fan_processor.pivot_table_resultfanJobs.to_csv()
                    st.download_button(
                        label="Download Fan Summary by Title CSV",
                        data=csv2,
                        file_name="fan_summary_by_title.csv",
                        mime="text/csv"
                    )

                    csv3 = fan_processor.pivot_table_fan.to_csv()
                    st.download_button(
                        label="Download Fan Summary by Location CSV",
                        data=csv3,
                        file_name="fan_summary_by_location.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Fan data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Fan analysis.")

        elif st.session_state.current_tab == 27:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Tank System Analysis")

            tank_processor = TankSystemProcessor()

            if ref_sheet is not None:
                try:
                    dftanks = pd.read_excel(ref_sheet, sheet_name='Tanks')
                    tank_processor.process_tank_data(data, dftanks)

                    st.subheader("Matched Tank Job Code Summary Table")
                    pivot_df = tank_processor.pivot_table_resulttanksJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_tanks = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_tanks,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped tank jobs found to display.")

                    st.subheader("Missing Job Codes from Tank Reference")
                    if not tank_processor.missingjobstankresult.empty:
                        st.dataframe(tank_processor.missingjobstankresult)
                    else:
                        st.success("All job codes are matched in the Tank reference sheet.")

                    # Downloads
                    csv1 = tank_processor.result_dftanks.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Tank Jobs CSV",
                        data=csv1,
                        file_name="matched_tank_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = tank_processor.pivot_table_resulttanksJobs.to_csv()
                    st.download_button(
                        label="Download Tank Summary Table CSV",
                        data=csv2,
                        file_name="tank_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Tank data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Tank analysis.")



        elif st.session_state.current_tab == 28:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Fresh Water Generator & Hydrophore System Analysis")

            fwg_processor = FWGSystemProcessor()

            if ref_sheet is not None:
                try:
                    dffwg = pd.read_excel(ref_sheet, sheet_name='FWG')
                    fwg_processor.process_fwg_data(data, dffwg)

                    st.subheader("Matched FWG & Hydrophore Job Code Summary Table")
                    pivot_df = fwg_processor.pivot_table_resultfwgJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_fwg = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_fwg,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped FWG or Hydrophore jobs found to display.")

                    st.subheader("Missing Job Codes from FWG Reference")
                    if not fwg_processor.missingjobsfwgresult.empty:
                        st.dataframe(fwg_processor.missingjobsfwgresult)
                    else:
                        st.success("All job codes are matched in the FWG reference sheet.")

                    # Downloads
                    csv1 = fwg_processor.result_dffwg.to_csv(index=False)
                    st.download_button(
                        label="Download Matched FWG Jobs CSV",
                        data=csv1,
                        file_name="matched_fwg_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = fwg_processor.pivot_table_resultfwgJobs.to_csv()
                    st.download_button(
                        label="Download FWG Summary Table CSV",
                        data=csv2,
                        file_name="fwg_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing FWG data: {str(e)}")
            else:
                st.warning("Reference sheet is required for FWG analysis.")


        elif st.session_state.current_tab == 29:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Workshop System Analysis")

            workshop_processor = WorkshopSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfworkshop = pd.read_excel(ref_sheet, sheet_name='Workshop')
                    workshop_processor.process_workshop_data(data, dfworkshop)

                    st.subheader("Matched Workshop Job Code Summary Table")
                    pivot_df = workshop_processor.pivot_table_resultworkshopJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_workshop = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_workshop,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped workshop jobs found to display.")

                    st.subheader("Missing Job Codes from Workshop Reference")
                    if not workshop_processor.missingjobsworkshopresult.empty:
                        st.dataframe(workshop_processor.missingjobsworkshopresult)
                    else:
                        st.success("All job codes are matched in the Workshop reference sheet.")

                    # Downloads
                    csv1 = workshop_processor.result_dfworkshop.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Workshop Jobs CSV",
                        data=csv1,
                        file_name="matched_workshop_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = workshop_processor.pivot_table_resultworkshopJobs.to_csv()
                    st.download_button(
                        label="Download Workshop Summary Table CSV",
                        data=csv2,
                        file_name="workshop_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Workshop data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Workshop analysis.")


        elif st.session_state.current_tab == 30:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Boiler System Analysis")

            boiler_processor = BoilerSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfboiler = pd.read_excel(ref_sheet, sheet_name='Boiler')
                    boiler_processor.process_boiler_data(data, dfboiler)

                    st.subheader("Matched Boiler Job Code Summary Table")
                    pivot_df = boiler_processor.pivot_table_resultboilerJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_boiler = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_boiler,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped boiler jobs found to display.")

                    st.subheader("Missing Job Codes from Boiler Reference")
                    if not boiler_processor.missingjobsboilerresult.empty:
                        st.dataframe(boiler_processor.missingjobsboilerresult)
                    else:
                        st.success("All job codes are matched in the Boiler reference sheet.")

                    # Downloads
                    csv1 = boiler_processor.result_dfboiler.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Boiler Jobs CSV",
                        data=csv1,
                        file_name="matched_boiler_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = boiler_processor.pivot_table_resultboilerJobs.to_csv()
                    st.download_button(
                        label="Download Boiler Summary Table CSV",
                        data=csv2,
                        file_name="boiler_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Boiler data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Boiler analysis.")



        elif st.session_state.current_tab == 31:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Miscellaneous System Analysis")

            misc_processor = MiscSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfmisc = pd.read_excel(ref_sheet, sheet_name='Misc')
                    misc_processor.process_misc_data(data, dfmisc)

                    # Matched Job Code Summary
                    st.subheader("Matched Misc Job Code Summary by Function")

                    pivot_df = misc_processor.pivot_table_resultmiscJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_misc = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_misc,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped miscellaneous jobs found to display.")

                    # Total Count Table
                    st.subheader("Total Misc Jobs by Title")
                    if misc_processor.styled_pivot_table_resultmiscJobstotal is not None:
                        st.markdown(
                            misc_processor.styled_pivot_table_resultmiscJobstotal.to_html(),
                            unsafe_allow_html=True
                        )
                    else:
                        st.dataframe(misc_processor.pivot_table_resultmiscJobstotal)

                    # Missing Job Codes
                    st.subheader("Missing Job Codes from Misc Reference")
                    if misc_processor.styled_missingmiscjobsresult is not None and not misc_processor.missingmiscjobsresult.empty:
                        st.markdown(
                            misc_processor.styled_missingmiscjobsresult.to_html(),
                            unsafe_allow_html=True
                        )
                    elif misc_processor.missingmiscjobsresult.empty:
                        st.success("All job codes are matched in the Misc reference sheet.")

                    # Downloads
                    csv1 = misc_processor.result_dfmisc.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Misc Jobs CSV",
                        data=csv1,
                        file_name="matched_misc_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = misc_processor.pivot_table_resultmiscJobs.to_csv()
                    st.download_button(
                        label="Download Misc Summary Table CSV",
                        data=csv2,
                        file_name="misc_summary_table.csv",
                        mime="text/csv"
                    )

                    csv3 = misc_processor.pivot_table_resultmiscJobstotal.to_csv()
                    st.download_button(
                        label="Download Misc Total Count CSV",
                        data=csv3,
                        file_name="misc_total_jobs_by_title.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Misc data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Misc analysis.")

        elif st.session_state.current_tab == 32:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Battery System Analysis")

            battery_processor = BatterySystemProcessor()

            if ref_sheet is not None:
                try:
                    dfbattery = pd.read_excel(ref_sheet, sheet_name='Battery')
                    battery_processor.process_battery_data(data, dfbattery)

                    st.subheader("Matched Battery Job Code Summary Table")
                    pivot_df = battery_processor.pivot_table_resultbatteryJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_battery = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_battery,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped battery jobs found to display.")

                    st.subheader("Missing Job Codes from Battery Reference")
                    if not battery_processor.missingjobsbatteryresult.empty:
                        st.dataframe(battery_processor.missingjobsbatteryresult)
                    else:
                        st.success("All job codes are matched in the Battery reference sheet.")

                    # Downloads
                    csv1 = battery_processor.result_dfbattery.to_csv(index=False)
                    st.download_button(
                        label="Download Matched Battery Jobs CSV",
                        data=csv1,
                        file_name="matched_battery_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = battery_processor.pivot_table_resultbatteryJobs.to_csv()
                    st.download_button(
                        label="Download Battery Summary Table CSV",
                        data=csv2,
                        file_name="battery_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing Battery data: {str(e)}")
            else:
                st.warning("Reference sheet is required for Battery analysis.")


        elif st.session_state.current_tab == 33:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Bow Thruster (BT) System Analysis")

            bt_processor = BTSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfBT = pd.read_excel(ref_sheet, sheet_name='BT')
                    bt_processor.process_bt_data(data, dfBT)

                    st.subheader("Matched BT Job Code Summary Table")
                    if bt_processor.pivot_table_resultBTJobs is not None and not bt_processor.pivot_table_resultBTJobs.empty:
                        styled_bt = bt_processor.pivot_table_resultBTJobs.style.applymap(color_binary_cells)

                        st.dataframe(
                            styled_bt,
                            use_container_width=True,
                            column_config={
                                bt_processor.pivot_table_resultBTJobs.index.name or bt_processor.pivot_table_resultBTJobs.columns[0]: st.column_config.TextColumn(
                                    label=bt_processor.pivot_table_resultBTJobs.index.name or bt_processor.pivot_table_resultBTJobs.columns[0],
                                    width="large"
                                )
                            }
                        )
                    else:
                        st.info("No mapped Bow Thruster jobs found to display.")

                    st.subheader("Missing Job Codes from BT Reference")
                    if not bt_processor.missingjobsBTresult.empty:
                        st.dataframe(bt_processor.missingjobsBTresult)
                    else:
                        st.success("All job codes are matched in the BT reference sheet.")

                    # Downloads
                    csv1 = bt_processor.result_dfBT.to_csv(index=False)
                    st.download_button(
                        label="Download Matched BT Jobs CSV",
                        data=csv1,
                        file_name="matched_bt_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = bt_processor.pivot_table_resultBTJobs.to_csv()
                    st.download_button(
                        label="Download BT Summary Table CSV",
                        data=csv2,
                        file_name="bt_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing BT data: {str(e)}")
            else:
                st.warning("Reference sheet is required for BT analysis.")


        elif st.session_state.current_tab == 34:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("LPSCR System Analysis")

            lpscr_processor = LPSCRSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfLPSCR = pd.read_excel(ref_sheet, sheet_name='LPSCRYANMAR')
                    lpscr_processor.process_lpscr_data(data, dfLPSCR)

                    st.subheader("Matched LPSCR Job Code Summary Table")

                    pivot_df = lpscr_processor.pivot_table_resultLPSCRJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_lpscr = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_lpscr,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped LP SCR jobs found to display.")

                    st.subheader("Missing Job Codes from LPSCR Reference")
                    if not lpscr_processor.missingjobsLPSCRresult.empty:
                        st.dataframe(lpscr_processor.missingjobsLPSCRresult)
                    else:
                        st.success("All job codes are matched in the LPSCR reference sheet.")

                    # Downloads
                    csv1 = lpscr_processor.result_dfLPSCR.to_csv(index=False)
                    st.download_button(
                        label="Download Matched LPSCR Jobs CSV",
                        data=csv1,
                        file_name="matched_lpscr_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = lpscr_processor.pivot_table_resultLPSCRJobs.to_csv()
                    st.download_button(
                        label="Download LPSCR Summary Table CSV",
                        data=csv2,
                        file_name="lpscr_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing LPSCR data: {str(e)}")
            else:
                st.warning("Reference sheet is required for LPSCR analysis.")

        elif st.session_state.current_tab == 35:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("HPSCR System Analysis")

            hpscr_processor = HPSCRSystemProcessor()

            if ref_sheet is not None:
                try:
                    dfHPSCR = pd.read_excel(ref_sheet, sheet_name='HPSCRHITACHI')
                    hpscr_processor.process_hpscr_data(data, dfHPSCR)

                    st.subheader("Matched HPSCR Job Code Summary Table")
                    pivot_df = hpscr_processor.pivot_table_resultHPSCRJobs  # ✅ correct attribute

                    if pivot_df is not None and not pivot_df.empty:
                        styled_hpscr = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_hpscr,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped HP SCR jobs found to display.")

                    st.subheader("Missing Job Codes from HPSCR Reference")
                    if not hpscr_processor.missingjobsHPSCRresult.empty:
                        st.dataframe(hpscr_processor.missingjobsHPSCRresult)
                    else:
                        st.success("All job codes are matched in the HPSCR reference sheet.")

                    # Downloads
                    csv1 = hpscr_processor.result_dfHPSCR.to_csv(index=False)
                    st.download_button(
                        label="Download Matched HPSCR Jobs CSV",
                        data=csv1,
                        file_name="matched_hpscr_jobs.csv",
                        mime="text/csv"
                    )

                    csv2 = hpscr_processor.pivot_table_resultHPSCRJobs.to_csv()
                    st.download_button(
                        label="Download HPSCR Summary Table CSV",
                        data=csv2,
                        file_name="hpscr_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"Error processing HPSCR data: {str(e)}")
            else:
                st.warning("Reference sheet is required for HPSCR analysis.")


        elif st.session_state.current_tab == 36:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("LSA Mapping Analysis")

            lsa_processor = LSAMappingProcessor()

            if ref_sheet is not None:
                try:
                    dflsa = pd.read_excel(ref_sheet, sheet_name='lsamapping')
                    lsa_processor.process_lsa_data(data, dflsa)

                    st.subheader("LSA Mapping Summary by Function")

                    pivot_df = lsa_processor.pivot_table_resultlsaJobs

                    if pivot_df is not None and not pivot_df.empty:
                        styled_lsa = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_lsa,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No mapped LSA job summary found.")

                    st.subheader("Total Mapped LSA Jobs by Title")

                    pivot_df = lsa_processor.pivot_table_resultlsaJobstotal

                    if pivot_df is not None and not pivot_df.empty:
                        styled_total_lsa = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name
                            if pivot_df.index.name
                            else pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_total_lsa,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No total mapped LSA job data found.")
                    st.subheader("Missing Job Codes from LSA Reference")
                    if not lsa_processor.missinglsajobsresult.empty:
                        st.dataframe(lsa_processor.missinglsajobsresult, use_container_width=True)
                    else:
                        st.success("All job codes are matched in the LSA reference sheet.")

                    # Download buttons
                    st.download_button(
                        label="Download Mapped LSA Jobs CSV",
                        data=lsa_processor.result_dflsa.to_csv(index=False),
                        file_name="mapped_lsa_jobs.csv",
                        mime="text/csv"
                    )
                    st.download_button(
                        label="Download LSA Summary Table CSV",
                        data=lsa_processor.pivot_table_resultlsaJobs.to_csv(),
                        file_name="lsa_summary_table.csv",
                        mime="text/csv"
                    )

                except Exception as e:
                    st.error(f"❌ Error processing LSA Mapping: {str(e)}")
            else:
                st.warning("Reference sheet is required for LSA Mapping analysis.")


        elif st.session_state.current_tab == 37:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("FFA Mapping Analysis")

            ffa_processor = FFAMappingProcessor()

            if ref_sheet is not None:
                try:
                    dfffa = pd.read_excel(ref_sheet, sheet_name='ffamapping')
                    ffa_processor.process_ffa_data(data, dfffa)

                    st.subheader("FFA Mapping Summary by Function")
                    pivot_df = ffa_processor.pivot_table_resultffaJobs

                    if pivot_df is not None and not pivot_df.empty:
                        styled_ffa = pivot_df.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df.index.name or pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_ffa,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No FFA function mapping summary data available.")

                    st.subheader("Total Mapped FFA Jobs by Title")
                    pivot_df_total = ffa_processor.pivot_table_resultffaJobstotal

                    if pivot_df_total is not None and not pivot_df_total.empty:
                        styled_total_ffa = pivot_df_total.style.applymap(color_binary_cells)

                        first_col = (
                            pivot_df_total.index.name or pivot_df_total.columns[0] if len(pivot_df_total.columns) > 0 else "Column 1"
                        )

                        st.dataframe(
                            styled_total_ffa,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No total mapped FFA job data available.")

                    st.subheader("Missing Job Codes from FFA Reference")
                    if not ffa_processor.missingffajobsresult.empty:
                        st.dataframe(ffa_processor.missingffajobsresult, use_container_width=True)
                    else:
                        st.success("All job codes are matched in the FFA reference sheet.")

                except Exception as e:
                    st.error(f"Error processing FFA Mapping: {str(e)}")
            else:
                st.warning("Reference sheet is required for FFA Mapping analysis.")


        elif st.session_state.current_tab == 38:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Inactive Mapping Analysis")

            inactive_processor = InactiveMappingProcessor()

            if ref_sheet is not None:
                try:
                    dfinactive = pd.read_excel(ref_sheet, sheet_name='inactivemapping')
                    inactive_processor.process_inactive_data(data, dfinactive)

                    st.subheader("Inactive Mapping Summary by Function")
                    pivot_df = inactive_processor.pivot_table_resultinactiveJobs

                    if pivot_df is not None and not pivot_df.empty:
                        styled_inactive = pivot_df.style.applymap(color_binary_cells)

                        first_col = pivot_df.index.name or pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"

                        st.dataframe(
                            styled_inactive,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No inactive function mapping summary data available.")

                    st.subheader("Total Mapped Inactive Jobs by Title")
                    pivot_df_total = inactive_processor.pivot_table_resultinactiveJobstotal

                    if pivot_df_total is not None and not pivot_df_total.empty:
                        styled_total_inactive = pivot_df_total.style.applymap(color_binary_cells)

                        first_col = pivot_df_total.index.name or pivot_df_total.columns[0] if len(pivot_df_total.columns) > 0 else "Column 1"

                        st.dataframe(
                            styled_total_inactive,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No total mapped inactive job data available.")

                    st.subheader("Missing Job Codes from Inactive Reference")
                    if not inactive_processor.missinginactivejobsresult.empty:
                        st.dataframe(inactive_processor.missinginactivejobsresult, use_container_width=True)
                    else:
                        st.success("All job codes are matched in the Inactive reference sheet.")

                except Exception as e:
                    st.error(f"Error processing Inactive Mapping: {str(e)}")
            else:
                st.warning("Reference sheet is required for Inactive Mapping analysis.")

        elif st.session_state.current_tab == 39:
            with st.container():
                col1, col2, col3 = st.columns([1, 1, 2])
                with col3:
                    st.header("Critical Jobs Mapping Analysis")

            critical_processor = CriticalJobsProcessor()

            if ref_sheet is not None:
                try:
                    dfcritical = pd.read_excel(ref_sheet, sheet_name='criticalmapping')
                    critical_processor.process_critical_data(data, dfcritical)

                    st.subheader("Critical Mapping Summary by Function")
                    pivot_df = critical_processor.pivot_table_resultcriticalJobs

                    if pivot_df is not None and not pivot_df.empty:
                        styled_critical = pivot_df.style.applymap(color_binary_cells)

                        first_col = pivot_df.index.name or pivot_df.columns[0] if len(pivot_df.columns) > 0 else "Column 1"

                        st.dataframe(
                            styled_critical,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No critical mapping summary data available.")

                    st.subheader("Total Mapped Critical Jobs by Title")
                    pivot_df_total = critical_processor.pivot_table_resultcriticalJobstotal

                    if pivot_df_total is not None and not pivot_df_total.empty:
                        styled_critical_total = pivot_df_total.style.applymap(color_binary_cells)

                        first_col = pivot_df_total.index.name or pivot_df_total.columns[0] if len(pivot_df_total.columns) > 0 else "Column 1"

                        st.dataframe(
                            styled_critical_total,
                            use_container_width=True,
                            column_config={
                                first_col: st.column_config.TextColumn(label=first_col, width="large")
                            }
                        )
                    else:
                        st.info("No total mapped critical job data available.")

                    st.subheader("Missing Job Codes from Critical Reference")
                    if not critical_processor.missingcriticaljobsresult.empty:
                        st.dataframe(critical_processor.missingcriticaljobsresult, use_container_width=True)
                    else:
                        st.success("All job codes are matched in the Critical reference sheet.")

                except Exception as e:
                    st.error(f"Error processing Critical Mapping: {str(e)}")
            else:
                st.warning("Reference sheet is required for Critical Mapping analysis.")

        elif st.session_state.current_tab == 40:
            st.header("QuickView Summary")

            if ref_sheet is not None:
                try:
                    # Load reference sheets
                    ref_sheets = pd.read_excel(ref_sheet, sheet_name=None)
                    dfML = ref_sheets.get('Machinery Location', pd.DataFrame())
                    dfCM = ref_sheets.get('Critical Machinery', pd.DataFrame())
                    dfVSM = ref_sheets.get('Vessel Specific Machinery', pd.DataFrame())

                    analyzer = QuickViewAnalyzer(data, dfML, dfCM, dfVSM)

                    ae_ref_pivot, ae_missing_jobs = ae_processor.process_reference_data(data, ref_sheet)
                    battery_processor.process_battery_data(data, ref_sheets.get("Battery", pd.DataFrame()))
                    boat_processor.process_boat_data(data, ref_sheets.get("Boats", pd.DataFrame()))
                    boiler_processor.process_boiler_data(data, ref_sheets.get("Boiler", pd.DataFrame()))
                    bridge_processor.process_bridge_data(data, ref_sheets.get("Bridge", pd.DataFrame()))
                    bt_processor.process_bt_data(data, ref_sheets.get("Bow Thruster", pd.DataFrame()))
                    bwts_missing_jobs = bwts_processor.process_reference_data(data, ref_sheet)

                    chs_processor = CargoHandlingSystemProcessor()
                    dfCargoHandling = pd.read_excel(ref_sheet, sheet_name="Cargohanding")
                    chs_processor.process_reference_data(data, ref_sheet)

                    cargopumping_processor.process_reference_data(data, ref_sheet)
                    cargoventing_processor.process_reference_data(data, ref_sheet)
                    compressor_processor.process_compressor_data(data, ref_sheets.get("Compressor", pd.DataFrame()))
                    crane_processor.process_crane_data(data, ref_sheets.get("Crane", pd.DataFrame()))
                    critical_processor.process_critical_data(data, ref_sheets.get("criticalmapping", pd.DataFrame()))
                    main_engine_data, *_ , missing_jobs, _ = process_engine_data(data, ref_sheet, engine_type)
                    ffamapping_processor.process_ffa_data(data, pd.read_excel(ref_sheet, sheet_name="ffamapping"))
                    fwg_processor.process_fwg_data(data, pd.read_excel(ref_sheet, sheet_name="FWG"))
                    missing_jobs_hatch = hatch_processor.process_reference_data(data, ref_sheet)
                    hpscr_processor.process_hpscr_data(data, pd.read_excel(ref_sheet, sheet_name="HPSCRHITACHI"))
                    inactive_processor.process_inactive_data(data, pd.read_excel(ref_sheet, sheet_name="inactivemapping"))
                    inertgas_processor.process_reference_data(data, ref_sheet)
                    ladder_processor.process_ladder_data(data, pd.read_excel(ref_sheet, sheet_name="Ladders"))
                    incin_processor.process_incin_data(data, pd.read_excel(ref_sheet, sheet_name="Incin"))
                    lpscr_processor.process_lpscr_data(data, pd.read_excel(ref_sheet, sheet_name="LPSCRYANMAR"))
                    lsamapping_processor.process_lsa_data(data, pd.read_excel(ref_sheet, sheet_name="lsamapping"))
                    misc_processor.process_misc_data(data, pd.read_excel(ref_sheet, sheet_name="Misc"))
                    mooring_processor.process_mooring_data(data, pd.read_excel(ref_sheet, sheet_name="Mooring"))
                    ows_processor.process_ows_data(data, pd.read_excel(ref_sheet, sheet_name="OWS"))
                    powerdist_processor.process_powerdist_data(data, pd.read_excel(ref_sheet, sheet_name="Powerdist"))
                    missingjobspurifierresult = purifier_processor.process_reference_data(data, ref_sheet)
                    refac_processor.process_refac_data(data, pd.read_excel(ref_sheet, sheet_name="Refac"))
                    steering_processor.process_steering_data(data, pd.read_excel(ref_sheet, sheet_name="Steering"))
                    stp_processor.process_stp_data(data, pd.read_excel(ref_sheet, sheet_name="STP"))
                    tank_processor.process_tank_data(data, pd.read_excel(ref_sheet, sheet_name="Tanks"))
                    workshop_processor.process_workshop_data(data, pd.read_excel(ref_sheet, sheet_name="Workshop"))

                    # Collect job count summaries
                    vesselname, totaljobs, criticaljobscount, total_missing_jobs, missing_jobs_df, missing_machinery_count = analyzer.get_basic_counts(
                        ae_missing_jobs=ae_missing_jobs,
                        battery_missing_jobs=battery_processor.missingjobsbatteryresult,
                        boat_missing_jobs=boat_processor.missingjobsBoatsresult,
                        boiler_missing_jobs=boiler_processor.missingjobsboilerresult,
                        bridge_missing_jobs=bridge_processor.missingjobsbridgeresult,
                        bt_missing_jobs=bt_processor.missingjobsBTresult,
                        bwts_missing_jobs=bwts_missing_jobs,
                        Cargo_Handling_System=chs_processor.missing_jobs_cargohandling.copy(),
                        Cargo_Pumping_System=cargopumping_processor.missingjobscargopumpingresult.copy(),
                        Cargo_Venting_System=cargoventing_processor.missing_jobs_cargovent.copy(),
                        compressor_missing_jobs=compressor_processor.missingjobsCompressorresult,
                        crane_missing_jobs=crane_processor.missingjobscraneresult,
                        Critical_Jobs=critical_processor.missingcriticaljobsresult.copy(),
                        Main_Engine=missing_jobs.copy(),
                        FFA_Mapping=ffamapping_processor.missingffajobsresult.copy(),
                        FWG_System=fwg_processor.missingjobsfwgresult.copy(),
                        Hatch_System=missing_jobs_hatch.copy(),
                        HPSCR_System=hpscr_processor.missingjobsHPSCRresult.copy(),
                        Inactive_Jobs=inactive_processor.missinginactivejobsresult.copy(),
                        Inert_Gas_System=inertgas_processor.missing_jobs_igsystem.copy(),
                        Ladder_System=ladder_processor.missingjobsLadderresult.copy(),
                        Incinerator_System=incin_processor.missingjobsIncinresult.copy(),
                        LPSCR_System=lpscr_processor.missingjobsLPSCRresult.copy(),
                        LSA_Mapping=lsamapping_processor.missinglsajobsresult.copy(),
                        Misc_Jobs=misc_processor.missingmiscjobsresult.copy(),
                        Mooring_System=mooring_processor.missingjobsMooringresult.copy(),
                        OWS_System=ows_processor.missingjobsOWSresult.copy(),
                        Power_Distribution_System=powerdist_processor.missingjobspowerdistresult.copy(),
                        Purifier_System=missingjobspurifierresult.copy(),
                        Refac_System=refac_processor.missingjobsrefacresult.copy(),
                        Steering_System=steering_processor.missingjobsSteeringresult.copy(),
                        STP_System=stp_processor.missingjobsSTPresult.copy(),
                        Tank_System=tank_processor.missingjobstankresult.copy(),
                        Workshop_System=workshop_processor.missingjobsworkshopresult.copy(),
                    )

                            # 🎯 Metrics Display - grouped layout
                    col1, col2, col3 = st.columns(3)
                    col1.metric("🛳️ Vessel", vesselname)
                    col2.metric("🧾 Total Jobs", totaljobs)
                    col3.metric("🚨 Critical Jobs", criticaljobscount)

                    col4, col5 = st.columns(2)
                    col4.metric("❌ Total Missing Jobs", total_missing_jobs)
                    col5.metric("🔧 Missing Machinery", missing_machinery_count)

                    # 📊 Add Pie Charts for Job and Machinery Comparison
                    import matplotlib.pyplot as plt

                    try:
                        def generate_pie_chart_with_values(labels, values, title):
                            fig, ax = plt.subplots(figsize=(4, 4))  # Smaller chart size
                            total = sum(values)

                            def format_label(pct):
                                absolute = int(round(pct / 100. * total))
                                return f"{pct:.1f}%\n({absolute})"

                            wedges, texts, autotexts = ax.pie(
                                values, labels=labels, autopct=lambda pct: format_label(pct),
                                startangle=90, textprops=dict(color="black")
                            )
                            ax.set_title(title, fontsize=10)
                            ax.axis('equal')
                            return fig

                        # 📊 Layout: Side-by-side columns
                        col_pie1, col_pie2 = st.columns(2)

                        with col_pie1:
                            st.subheader("📊 Jobs Summary")
                            job_pie = generate_pie_chart_with_values(
                                ["Total Jobs", "Missing Jobs"],
                                [totaljobs, total_missing_jobs],
                                "Total Jobs vs Missing Jobs"
                            )
                            st.pyplot(job_pie)

                        with col_pie2:
                            st.subheader("⚙️ Machinery Summary")
                            total_onboard_machinery = len(analyzer.df['Machinery Locationcopy'].dropna().astype(str).str.lower().str.strip().unique())
                            mach_pie = generate_pie_chart_with_values(
                                ["Present Machinery", "Missing Machinery"],
                                [total_onboard_machinery - missing_machinery_count, missing_machinery_count],
                                "Total Machinery vs Missing Machinery"
                            )
                            st.pyplot(mach_pie)

                    except Exception as pie_err:
                        st.warning(f"⚠️ Unable to render pie charts: {pie_err}")

                    # 📋 Job Source Summary and CMS Code Count in expandable panel
                    with st.expander("📋 View Job Source Breakdown"):
                        try:
                            jobsource_summary = analyzer.generate_jobsource_summary()
                            st.dataframe(jobsource_summary, use_container_width=True)
                            st.metric("🔢 CMS Code Entries", analyzer.cms_code_count)
                        except Exception as e:
                            st.warning(f"Job Source Summary unavailable: {e}")

                    # 📊 Missing Jobs Chart (Styled)
                    if not missing_jobs_df.empty:
                        st.subheader("📊 Missing Jobs by Machinery System")
                        import matplotlib.pyplot as plt
                        fig, ax = plt.subplots(figsize=(12, 6))
                        bars = ax.bar(missing_jobs_df["Machinery System"], missing_jobs_df["Missing Jobs Count"], color='#5DADE2')

                        for bar in bars:
                            height = bar.get_height()
                            ax.text(bar.get_x() + bar.get_width() / 2, height + 0.5, str(int(height)), ha='center', va='bottom', fontsize=9)

                        ax.set_title("📊 Missing Jobs by Machinery System", fontsize=14, weight='bold')
                        ax.set_xlabel("Machinery System", fontsize=12)
                        ax.set_ylabel("Missing Jobs Count", fontsize=12)
                        ax.set_xticks(range(len(missing_jobs_df["Machinery System"])))
                        ax.set_xticklabels(missing_jobs_df["Machinery System"], rotation=45, ha='right')
                        ax.grid(axis='y', linestyle='--', alpha=0.5)
                        st.pyplot(fig)

                    # 📁 Display Missing Jobs Table
                    st.subheader("🗂 Missing Jobs Summary Table")
                    st.dataframe(missing_jobs_df, use_container_width=True)

                except Exception as e:
                    st.error(f"Error in QuickView Summary: {str(e)}")
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        st.exception(e)
else:
    st.info("Please upload a data file to begin analysis.")

    # Show sample data format guide
    with st.expander("Data Format Guide"):
        st.markdown("""
        ### Required Columns
        Your data should contain the following key columns:
        - **Vessel**: Vessel name
        - **Job Code**: Maintenance task code
        - **Frequency**: Maintenance frequency
        - **Calculated Due Date**: When the maintenance is due
        - **Machinery Location**: Where the equipment is located
        - **Sub Component Location**: Component details

        ### Engine Types
        The application supports analysis for different engine types:
        - **Normal Main Engine**: Standard main engine configuration
        - **MAN ME-C and ME-B Engine**: Electronic controlled engines
        - **RT Flex Engine**: Common rail engines
        - **RTA Engine**: Mechanical engines
        - **UEC Engine**: Mitsubishi engines
        - **WINGD Engine**: WinGD X series engines

        ### Reference Sheet
        For comprehensive analysis, upload a reference sheet with the same column structure that contains expected maintenance data.
        """)
