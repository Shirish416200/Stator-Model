from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Iterable

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

APP_TITLE = "Ampere-Turn Engineering Calculator"
DEFAULT_WORKBOOK = "Amp_turn_input_sheet.xlsx"
INPUT_ROWS = [8, 9, 10, 11, 14, 15, 16, 17, 18, 21, 22, 23, 24, 25, 28]
INPUT_RANGE = INPUT_ROWS
OUTPUT_CELLS = {
    "Estimated Ampere-Turns": "C115",
    "Number of Turns in Lead Wire": "C92",
}

st.set_page_config(page_title=APP_TITLE, page_icon="⚡", layout="wide")
st.markdown(
    """
<style>
.block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
[data-testid="stMetric"] {background:#f7f9fc; border:1px solid #dfe6ef; padding:16px; border-radius:14px;}
.input-card {background:#fff4f8; border:1px solid #f3b7cf; border-radius:14px; padding:14px 16px; margin-bottom:12px;}
.small-note {color:#5b6573; font-size:.9rem;}
</style>
""",
    unsafe_allow_html=True,
)


def clean_label(value, fallback: str) -> str:
    if value is None:
        return fallback
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or fallback


def pick_sheet(workbook: openpyxl.Workbook):
    for ws in workbook.worksheets:
        if any(ws[f"C{r}"].value is not None for r in INPUT_RANGE):
            return ws
    return workbook.active


def label_for_row(ws, row: int) -> str:
    # Prefer the nearest descriptive cell to the left of column C.
    for col in (2, 1):
        value = ws.cell(row=row, column=col).value
        if value not in (None, ""):
            return clean_label(value, f"Input C{row}")
    return f"Input C{row}"


def numeric_default(value) -> float:
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    return 0.0


def locate_soffice() -> str:
    candidates = [
        shutil.which("libreoffice"),
        shutil.which("soffice"),
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
    ]
    for item in candidates:
        if item and Path(item).exists():
            return item
    raise RuntimeError(
        "LibreOffice was not found on the server. Install LibreOffice so the existing Excel formulas can be recalculated."
    )


def recalculate_with_libreoffice(source: Path, output_dir: Path) -> Path:
    soffice = locate_soffice()
    profile_dir = output_dir / "lo-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice,
        "--headless",
        f"-env:UserInstallation=file://{profile_dir}",
        "--convert-to",
        "xlsx",
        "--outdir",
        str(output_dir),
        str(source),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    recalculated = output_dir / source.name
    if result.returncode != 0 or not recalculated.exists():
        details = (result.stderr or result.stdout or "Unknown LibreOffice error").strip()
        raise RuntimeError(f"Workbook recalculation failed: {details}")
    return recalculated


def write_inputs_and_recalculate(workbook_bytes: bytes, sheet_name: str, values: dict[int, float]) -> tuple[bytes, openpyxl.Workbook]:
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        original = root / DEFAULT_WORKBOOK
        original.write_bytes(workbook_bytes)

        wb = openpyxl.load_workbook(original, data_only=False)
        ws = wb[sheet_name]
        for row, value in values.items():
            ws[f"C{row}"] = float(value)
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"

        staged = root / "input_for_recalc.xlsx"
        wb.save(staged)
        out = root / "recalculated"
        out.mkdir()
        calculated_path = recalculate_with_libreoffice(staged, out)
        calculated_bytes = calculated_path.read_bytes()
        calculated_wb = openpyxl.load_workbook(calculated_path, data_only=True)
        return calculated_bytes, calculated_wb


def find_relevant_outputs(ws, excluded: Iterable[str]) -> list[tuple[str, str, float]]:
    excluded_set = set(excluded)
    found: list[tuple[str, str, float]] = []
    terms = ("ampere", "amp turn", "amp-turn", "ampturn", "turns")
    for row in range(1, ws.max_row + 1):
        labels = [ws.cell(row, c).value for c in range(1, min(ws.max_column, 6) + 1)]
        text = " ".join(str(v) for v in labels if isinstance(v, str)).lower()
        if not any(term in text for term in terms):
            continue
        label = next((clean_label(v, "") for v in labels if isinstance(v, str) and v.strip()), f"Row {row}")
        for col in range(2, min(ws.max_column, 12) + 1):
            cell = ws.cell(row, col)
            if cell.coordinate in excluded_set:
                continue
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                found.append((label, cell.coordinate, float(cell.value)))
                break
    # Keep the dashboard useful without flooding it.
    unique = []
    seen = set()
    for item in found:
        key = (item[0], item[1])
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique[:8]


def parse_voltage(text: str) -> int | None:
    normalized = text.lower().replace("volts", "v").replace("volt", "v")
    if re.search(r"(^|\D)9\s*v($|\D)", normalized):
        return 9
    if re.search(r"(^|\D)16\s*v($|\D)", normalized):
        return 16
    return None


def extract_temperature_curves(ws) -> pd.DataFrame:
    """Read the actual Amp-Turn table from Calculations!C140:O146."""
    temperatures = [ws.cell(140, col).value for col in range(3, 16)]
    voltage_rows = {}
    for row in range(141, 147):
        voltage = ws.cell(row, 2).value
        if voltage in (9, 16):
            voltage_rows[int(voltage)] = row

    if 9 not in voltage_rows or 16 not in voltage_rows:
        return pd.DataFrame(columns=["Temperature (°C)", "9 V", "16 V"])

    records = []
    for index, temperature in enumerate(temperatures, start=3):
        amp9 = ws.cell(voltage_rows[9], index).value
        amp16 = ws.cell(voltage_rows[16], index).value
        if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in (temperature, amp9, amp16)):
            # The requested plot range ends at 150°C. The sheet also contains 160°C and 180°C.
            if -40 <= float(temperature) <= 150:
                records.append({
                    "Temperature (°C)": float(temperature),
                    "9 V": float(amp9),
                    "16 V": float(amp16),
                })
    return (pd.DataFrame(records)
            .sort_values("Temperature (°C)")
            .drop_duplicates("Temperature (°C)"))


def workbook_source() -> tuple[bytes | None, str]:
    bundled = Path(__file__).with_name(DEFAULT_WORKBOOK)
    uploaded = st.sidebar.file_uploader("Excel calculation workbook", type=["xlsx"])
    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name
    if bundled.exists():
        return bundled.read_bytes(), bundled.name
    return None, DEFAULT_WORKBOOK


st.title("⚡ Ampere-Turn Engineering Calculator")
st.caption("Excel-powered calculator with preserved formulas, dynamic results, and temperature curves at 9 V and 16 V.")

workbook_bytes, workbook_name = workbook_source()
if workbook_bytes is None:
    st.warning(
        f"Add **{DEFAULT_WORKBOOK}** beside `app.py`, or upload it from the sidebar. "
        "The workbook is required because all engineering results are calculated from its existing formulas."
    )
    st.stop()

try:
    source_wb = openpyxl.load_workbook(tempfile.SpooledTemporaryFile(), data_only=False)
except Exception:
    # Load from bytes without saving user content permanently.
    import io
    source_wb = openpyxl.load_workbook(io.BytesIO(workbook_bytes), data_only=False)

source_ws = pick_sheet(source_wb)
st.sidebar.success(f"Workbook: {workbook_name}")
st.sidebar.caption(f"Calculation sheet: {source_ws.title}")

left, right = st.columns([0.95, 1.35], gap="large")
with left:
    st.subheader("1. User Inputs")
    st.markdown('<div class="input-card">The pink input cells from C8:C28 are shown below. Numeric values only.</div>', unsafe_allow_html=True)
    entered = {}
    with st.form("engineering_inputs"):
        for row in INPUT_RANGE:
            cell = source_ws[f"C{row}"]
            label = label_for_row(source_ws, row)
            entered[row] = st.number_input(
                f"{label}  ·  C{row}",
                value=numeric_default(cell.value),
                format="%.6g",
                key=f"input_{row}",
            )
        calculate = st.form_submit_button("Calculate", type="primary", use_container_width=True)

if calculate:
    try:
        with st.spinner("Updating workbook and recalculating Excel formulas..."):
            result_bytes, result_wb = write_inputs_and_recalculate(workbook_bytes, source_ws.title, entered)
            result_ws = result_wb[source_ws.title]
            st.session_state["result_bytes"] = result_bytes
            st.session_state["result_sheet"] = source_ws.title
            st.session_state["outputs"] = {
                name: result_ws[address].value for name, address in OUTPUT_CELLS.items()
            }
            st.session_state["additional"] = find_relevant_outputs(result_ws, OUTPUT_CELLS.values())
            st.session_state["curves"] = extract_temperature_curves(result_ws)
    except Exception as exc:
        st.error(str(exc))

with right:
    st.subheader("2. Calculation Results")
    if "outputs" not in st.session_state:
        st.info("Enter the required values and select **Calculate**.")
    else:
        c1, c2 = st.columns(2)
        output_items = list(st.session_state["outputs"].items())
        for col, (name, value) in zip((c1, c2), output_items):
            display = f"{value:,.4g}" if isinstance(value, (int, float)) else "Not available"
            col.metric(name, display)

        additional = st.session_state.get("additional", [])
        if additional:
            st.markdown("#### Other workbook estimates")
            additional_df = pd.DataFrame(additional, columns=["Parameter", "Cell", "Calculated value"])
            st.dataframe(additional_df, hide_index=True, use_container_width=True)

        st.download_button(
            "Download recalculated workbook",
            data=st.session_state["result_bytes"],
            file_name="Amp_turn_calculated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.subheader("3. Ampere-Turns vs Temperature")
    curves = st.session_state.get("curves")
    if isinstance(curves, pd.DataFrame) and not curves.empty:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=curves["Temperature (°C)"], y=curves["9 V"], mode="lines+markers", name="9 V", line=dict(color="#2563EB", width=3)))
        fig.add_trace(go.Scatter(x=curves["Temperature (°C)"], y=curves["16 V"], mode="lines+markers", name="16 V", line=dict(color="#E11D48", width=3)))
        fig.update_layout(
            title="Ampere-Turns vs Temperature at 9 V and 16 V",
            xaxis_title="Temperature (°C)",
            yaxis_title="Ampere-Turns",
            xaxis=dict(range=[-40, 150], gridcolor="#E5E7EB"),
            yaxis=dict(gridcolor="#E5E7EB"),
            template="plotly_white",
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=20, r=20, t=80, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(curves, hide_index=True, use_container_width=True)
    elif "outputs" in st.session_state:
        st.warning(
            "The workbook recalculated, but no table with Temperature, 9 V, and 16 V headers was detected. "
            "Keep those headers in the Excel estimation table so the graph uses the actual calculated values."
        )
