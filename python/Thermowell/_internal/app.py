import re
import math as m
from io import BytesIO
import sys
from pathlib import Path

import tempfile
import os
import pandas as pd
import pythoncom
import streamlit as st
import win32com.client
from openpyxl import load_workbook


st.set_page_config(
    page_title="Thermowell template generation from UNISIM",
    layout="wide"
)
# ============================================================
# UniSim Functions
# ============================================================

def get_unisim():
    """
    Create UniSim COM connection.
    """
    pythoncom.CoInitialize()

    try:
        app = win32com.client.Dispatch(
            "UniSimDesign.Application"
        )
        return app

    except Exception:
        pythoncom.CoUninitialize()
        raise


def parse_mapping(text):
    """
    Accepts:
        TAG1, STREAM1
        TAG2, STREAM2

    or pasted from Excel:
        TAG1    STREAM1
        TAG2    STREAM2
    """

    mappings = []

    for line in text.splitlines():

        line = line.strip()

        if not line:
            continue

        parts = [
            p.strip()
            for p in re.split(r'[,\t]+', line)
        ]

        if len(parts) < 2:
            continue

        mappings.append(
            {
                "Tag": parts[0],
                "Stream": parts[1]
            }
        )

    return mappings

def safe_float(value, default=None):
    try:
        number = float(value)
        return number if m.isfinite(number) else default
    except (TypeError, ValueError, OverflowError):
        return default


def get_phase_value(phase, property_name, default=None):
    try:
        return safe_float(
            getattr(phase, property_name),
            default
        )
    except Exception:
        return default


def calculate_liquid_viscosity(fluid):
    light_flow = get_phase_value(
        fluid.LightLiquidPhase,
        "MolarFlowValue"
    )
    heavy_flow = get_phase_value(
        fluid.HeavyLiquidPhase,
        "MolarFlowValue"
    )

    if light_flow is None:
        x_oil = 0.0
    elif heavy_flow is None:
        x_oil = 1.0 if light_flow > 0 else 0.0
    else:
        total_flow = light_flow + heavy_flow
        x_oil = light_flow / total_flow if total_flow > 0 else 0.0

    x_oil = max(0.0, min(1.0, x_oil))
    x_aq = 1.0 - x_oil

    mu_oil = get_phase_value(
        fluid.LightLiquidPhase,
        "ViscosityValue"
    )
    mu_aq = get_phase_value(
        fluid.HeavyLiquidPhase,
        "ViscosityValue"
    )

    if mu_oil is not None and mu_oil > 0 and mu_aq is not None and mu_aq > 0:
        mu_liquid = m.exp(
            x_oil * m.log(mu_oil)
            + x_aq * m.log(mu_aq)
        )
    elif mu_oil is not None and mu_oil > 0:
        mu_liquid = mu_oil
    elif mu_aq is not None and mu_aq > 0:
        mu_liquid = mu_aq
    else:
        mu_liquid = None

    return mu_liquid


def calculate_two_phase_viscosity(stream, fluid):
    mu_liquid = calculate_liquid_viscosity(fluid)

    mu_vapor = get_phase_value(
        fluid.VapourPhase,
        "ViscosityValue"
    )

    vapor_fraction = safe_float(
        stream.VapourFractionValue,
        default=0.0
    )
    vapor_fraction = max(0.0, min(1.0, vapor_fraction))

    if mu_liquid is None:
        return mu_vapor

    if mu_vapor is None:
        return mu_liquid

    numerator = (2*mu_liquid + mu_vapor - 2*(mu_liquid - mu_vapor)*vapor_fraction)
    denominator = (2*mu_liquid + mu_vapor + (mu_liquid - mu_vapor)*vapor_fraction)

    if abs(denominator) < 1e-12:
        return mu_liquid

    return mu_liquid * numerator / denominator

def get_multiple_streams(
    case_path,
    mappings,
    process_condition="",
    gas_limit_pct=5.0
):
    app = get_unisim()
    case = None

    gas_limit = safe_float(
        gas_limit_pct,
        default=5.0
    ) / 100.0

    gas_limit = max(
        0.0,
        min(1.0, gas_limit)
    )

    try:
        case = app.SimulationCases.Open(
            case_path
        )

        results = []

        for item in mappings:

            tag = str(
                item.get("Tag", "")
            ).strip()

            stream_name = str(
                item.get("Stream", "")
            ).strip()

            try:
                stream = case.Flowsheet.MaterialStreams.Item(
                    stream_name
                )

                liquid_fraction = safe_float(
                    stream.LiquidFractionValue
                )

                if liquid_fraction is None:
                    raise ValueError(
                        "LiquidFractionValue is unavailable."
                    )

                liquid_fraction = max(
                    0.0,
                    min(1.0, liquid_fraction)
                )

                fluid = stream.GetFluid()

                if liquid_fraction <= 1e-12:
                    fluid_description = "Gas"
                    physical_state = "Gas"

                    dynamic_viscosity = get_phase_value(
                        fluid.VapourPhase,
                        "ViscosityValue",
                        default=safe_float(stream.ViscosityValue)
                    )

                elif liquid_fraction <= gas_limit:
                    gas_percentage = (
                        1.0 - liquid_fraction
                    ) * 100.0

                    fluid_description = (
                        f"Gas ({gas_percentage:.0f}%)"
                    )
                    physical_state = "Gas"

                    dynamic_viscosity = (
                        calculate_two_phase_viscosity(
                            stream,
                            fluid
                        )
                    )

                elif liquid_fraction < 1.0 - 1e-12:
                    liquid_percentage = (
                        liquid_fraction * 100.0
                    )

                    fluid_description = (
                        f"Liquid ({liquid_percentage:.0f}%)"
                    )
                    physical_state = "Liquid"

                    dynamic_viscosity = (
                        calculate_two_phase_viscosity(
                            stream,
                            fluid
                        )
                    )

                else:
                    fluid_description = "Liquid"
                    physical_state = "Liquid"

                    dynamic_viscosity = (
                        calculate_liquid_viscosity(fluid)
                    )

                    if dynamic_viscosity is None:
                        dynamic_viscosity = safe_float(
                            stream.ViscosityValue
                        )

                temperature = safe_float(
                    stream.TemperatureValue
                )

                pressure_kpa_abs = safe_float(
                    stream.PressureValue
                )

                density = safe_float(
                    stream.MassDensityValue
                )

                volume_flow_m3s = safe_float(
                    stream.ActualVolumeFlowValue
                )

                pressure_barg = (
                    (pressure_kpa_abs - 101.325) / 100.0
                    if pressure_kpa_abs is not None
                    else ""
                )

                volume_flow_m3h = (
                    volume_flow_m3s * 3600.0
                    if volume_flow_m3s is not None
                    else ""
                )

                results.append(
                    {
                        "Tag": tag,
                        "Process Condition": process_condition,
                        "Fluid Description": fluid_description,
                        "Fluid Physical State": physical_state,
                        "Erosive Application": "No",
                        "Dynamic Viscosity": (
                            dynamic_viscosity
                            if dynamic_viscosity is not None
                            else ""
                        ),
                        "Temperature": (
                            temperature
                            if temperature is not None
                            else ""
                        ),
                        "Pressure": pressure_barg,
                        "Density": (
                            density
                            if density is not None
                            else ""
                        ),
                        "Volumetric Flow Rate": volume_flow_m3h,
                        "Status": "OK"
                    }
                )

            except Exception as error:
                results.append(
                    {
                        "Tag": tag,
                        "Process Condition": process_condition,
                        "Fluid Description": "",
                        "Fluid Physical State": "",
                        "Erosive Application": "",
                        "Dynamic Viscosity": "",
                        "Temperature": "",
                        "Pressure": "",
                        "Density": "",
                        "Volumetric Flow Rate": "",
                        "Status": f"ERROR: {error}"
                    }
                )

        return pd.DataFrame(results)

    finally:

        if case is not None:
            case.Close()

        pythoncom.CoUninitialize()

# ============================================================
# Excel Template Function
# ============================================================

def create_tda_import(
    df,
    template_path=None
):

    if template_path is None:
        template_path = resource_path(
            "TDA_ImportTemplate.xlsx"
        )

    if not Path(template_path).exists():
        raise FileNotFoundError(
            f"Template not found: {template_path}"
        )

    wb = load_workbook(
        template_path,
        keep_vba=False,
        data_only=False
    )

    ws = wb["WFC Data"]

    first_data_row = 19

    for _, row in df.iterrows():

        tag = str(row["Tag"]).strip()

        # Do not export rows where the UniSim stream failed
        if row["Status"] != "OK":
            continue

        target_row = None

        # Search for an existing matching tag
        for excel_row in range(
            first_data_row,
            ws.max_row + 1
        ):

            cell_value = ws[f"A{excel_row}"].value

            if cell_value is None:
                continue

            existing_tag = str(cell_value).strip()

            if existing_tag.casefold() == tag.casefold():
                target_row = excel_row
                break

        # If the tag does not exist, find the first empty tag row
        if target_row is None:

            for excel_row in range(
                first_data_row,
                ws.max_row + 1
            ):

                cell_value = ws[f"A{excel_row}"].value

                if cell_value is None or str(cell_value).strip() == "":
                    target_row = excel_row
                    break

        # If every existing row is occupied, append after max_row
        if target_row is None:
            target_row = ws.max_row + 1

        # Write the imported data
        ws[f"A{target_row}"] = row["Tag"]
        ws[f"B{target_row}"] = row["Process Condition"]
        ws[f"C{target_row}"] = row["Fluid Description"]
        ws[f"D{target_row}"] = row["Fluid Physical State"]
        ws[f"E{target_row}"] = row["Erosive Application"]
        ws[f"F{target_row}"] = row["Dynamic Viscosity"]
        ws[f"G{target_row}"] = row["Temperature"]
        ws[f"H{target_row}"] = row["Pressure"]
        ws[f"I{target_row}"] = row["Density"]
        ws[f"M{target_row}"] = row["Volumetric Flow Rate"]

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    return output

def resource_path(relative_path):
    """
    Get absolute path to bundled resource.
    Works both in development and PyInstaller.
    """
    if getattr(sys, "frozen", False):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(__file__).parent

    return base_path / relative_path

# ============================================================
# UI
# ============================================================
def main():
    
    if "step" not in st.session_state:
        st.session_state.step = 1

    st.title("Emerson Thermowell Template Generator")

    st.progress(st.session_state.step / 3)

    st.caption(
        f"Step {st.session_state.step} of 3"
    )

    if st.session_state.step == 1:

        st.subheader("Step 1: Select UniSim File")

        uploaded_case = st.file_uploader(
            "UniSim File",
            type=["usc"]
        )

        process_condition = st.text_input(
            "Case Number (optional)",
            ""
        )

        col1, col2 = st.columns(2)

        with col2:
            if st.button("Next"):

                if uploaded_case is None:
                    st.error(
                        "Please select a UniSim case file."
                    )

                else:
                    st.session_state.uploaded_case = uploaded_case
                    st.session_state.process_condition = (
                        process_condition
                    )
                    st.session_state.step = 2
                    st.rerun()

    elif st.session_state.step == 2:

        col1, col2 = st.columns(2)

        with col1:
            st.subheader(
                "Step 2: Stream Configuration"
            )   
        
        with col2:
            left_spacer, right_content = st.columns([1, 1])
            with right_content:
                with st.popover("ℹ️ Note on liquid fraction limit and viscosity"):

                    st.markdown("""
                    ### Liquid fraction limit

                    Streams with liquid fraction less than or equal to the limit will be classified as Gas in the template.


                    ### Viscosity calculations
                    For pure gas or liquid the UNISIM value is used directly.
                    If multiple phases are present (gas-liquid, liquid-liquid, gas-liquid-liquid, ...), 
                    the viscosity is calculated with the following set of equations independently of
                    clasification in template (liquid fraction limit).

                    For two-phase (gas and liquid) flow, Awad and Muzychka model [1] is used. 
                    
                    """)

                    st.latex(r"""
                    \mu_m = \mu_l \frac{2\mu_l + \mu_g - 2(\mu_l-\mu_g)w}{2\mu_l+\mu_g+(\mu_l - \mu_g)w}
                    """)

                    st.markdown(r"""
                    Where $w$ is mass fraction of gas in the mixture and $\mu$ is dynamic viscosity.

                    If multiple liquid phases is present, 

                    """)

                    st.latex(r"""
                    \mu_l = exp(\sum_i x_i ln(\mu_i))
                    """)

                    st.markdown(r"""
                    Where $x_i$ is mole fraction of the given phase in the liquid.

                    [1] Ghajar, A. J., & Bhagwat, S. M. (2013). 
                    Effect of Void Fraction and Two-Phase Dynamic Viscosity Models on Prediction of 
                    Hydrostatic and Frictional Pressure Drop in Vertical Upward Gas–Liquid Two-Phase Flow. 
                    Heat Transfer Engineering, 34(13), 1044–1059. https://doi.org/10.1080/01457632.2013.763541

                    """)

        with col1:
            gas_limit_pct = st.number_input(
                "Liquid Fraction Limit (%)",
                min_value=0.0,
                max_value=100.0,
                value=5.0,
                step=1.0
            )

        mapping_text = st.text_area(
            "Paste Tag / Stream mappings",
            height=250,
            placeholder="""TAG, UNISIM STREAM

    for example:
    20TW001, 20
    20TW002, 21
    20TW003, 22
    """
        )

        col1, col2 = st.columns(2)

        with col1:
            spacer, button_col = st.columns([4, 1])
            with button_col:
                if st.button("Back"):
                    st.session_state.step = 1
                    st.rerun()

        with col2:

            if st.button("Next"):

                mappings = parse_mapping(
                    mapping_text
                )

                if not mappings:

                    st.error(
                        "No valid mappings found."
                    )

                else:

                    st.session_state.gas_limit_pct = (
                        gas_limit_pct
                    )

                    st.session_state.mapping_text = (
                        mapping_text
                    )

                    st.session_state.step = 3
                    st.rerun()

    elif st.session_state.step == 3:

        st.subheader(
            "Step 3: Template Selection"
        )

        st.caption("Upload a pre-filled template. "
                "If no file is uploaded, "
                "a new template will be used.")

        st.caption("If a tag in the pre-filled match the mapping, "
                "UNISIM data will be replaced."  
                "If a tag in the mapping does not match, it will be appended to the sheet.")

        uploaded_template = st.file_uploader(
            "Excel Template (optional)",
            type=["xlsx"],
        )

        col1, col2 = st.columns(2)

        with col1:
            spacer, button_col = st.columns([4, 1])
            with button_col:
                if st.button("Back"):
                    st.session_state.step = 2
                    st.rerun()

        with col2:

            run_generation = st.button(
                "Generate Template"
            )

        if run_generation:

            uploaded_case = (
                st.session_state.uploaded_case
            )

            process_condition = (
                st.session_state.process_condition
            )

            gas_limit_pct = (
                st.session_state.gas_limit_pct
            )

            mappings = parse_mapping(
                st.session_state.mapping_text
            )

            template_path = None
            case_path = None

            try:

                with tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".usc"
                ) as temp_file:

                    temp_file.write(
                        uploaded_case.getbuffer()
                    )

                    case_path = temp_file.name

                if uploaded_template is not None:

                    with tempfile.NamedTemporaryFile(
                        delete=False,
                        suffix=".xlsx"
                    ) as temp_template:

                        temp_template.write(
                            uploaded_template.getbuffer()
                        )

                        template_path = (
                            temp_template.name
                        )

                with st.spinner(
                    "Opening UniSim and reading stream data..."
                ):

                    df = get_multiple_streams(
                        case_path,
                        mappings,
                        process_condition,
                        gas_limit_pct
                    )

                st.success(
                    f"Processed {len(df)} mappings."
                )

                st.dataframe(
                    df,
                    use_container_width=True
                )

                with st.spinner(
                    "Generating Thermowell template file..."
                ):

                    tda_file = create_tda_import(
                        df,
                        template_path=template_path
                    )

                output_name = (
                    Path(
                        uploaded_case.name
                    ).stem
                    + "_TDA_Import.xlsx"
                )

                st.download_button(
                    label="Download Thermowell Template",
                    data=tda_file,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:

                st.exception(e)

            finally:

                if (
                    case_path
                    and Path(case_path).exists()
                ):
                    os.remove(case_path)

                if (
                    template_path
                    and Path(template_path).exists()
                ):
                    os.remove(template_path)

if __name__ == "__main__":
    main()