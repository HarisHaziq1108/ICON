"""Shared profiler and economics calculations and workbook access."""

import calendar
import hashlib
import json
import math
import os
import re
import tempfile
import threading
import unicodedata
import zipfile
from datetime import datetime
from functools import lru_cache

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


VARIABLES = [
    ("Oil_Sales", "Oil Sales (Mstb/d)"),
    ("Gas_Sales", "Gas Sales (MMscf/d)"),
    ("Condy_Sales", "Condy Sales (Mstb/d)"),
    ("Exploration_Capex", "Exploration Capex (USD mill)"),
    ("Appraisal_Wells", "Appraisal Wells (USD mill)"),
    ("Jacket", "Jacket (USD mill)"),
    ("MachineryandPipelines", "Machinery & Pipelines (USD mill)"),
    ("Drilling", "Drilling (USD mill)"),
    ("Fixed_Opex", "Fixed Opex (USD mill)"),
    ("Variable_Opex", "Variable Opex (USD mill)"),
    ("lease", "Lease (USD mill)"),
    ("Well_PnA", "Well P&A (USD mill)"),
    ("Abex_Cess", "Abex Cess (USD mill)"),
    ("Liquid_Capex", "Liquid Capex for Export Duty (USD mill)"),
    ("Liquid_Opex", "Liquid Opex for Export Duty (USD mill)"),
    ("Liquid_Abex", "Liquid Abex for Export Duty (USD mill)"),
    ("Project_NCF", "Project Net Cash Flow (USD mill)"),
]

YEARS = list(range(2020, 2061))

MIN_YEAR = min(YEARS)

MAX_YEAR = max(YEARS)

PROFILE_TYPES = {"notional", "complete"}
PROJECT_INFO_SHEET = "Project Information"

PRICE_SCENARIOS = {"base", "low", "high", "custom"}

PRICE_SCENARIO_LABELS = {
    "base": "Base KPBI FY2027",
    "low": "Low KPBI FY2027",
    "high": "High KPBI FY2027",
    "custom": "Custom Price",
}

WORKBOOK_ROYALTY_RATE = 0.10

WORKBOOK_EXPORT_DUTY_RATE = 0.10

WORKBOOK_PITA_RATE = 0.38

WORKBOOK_INITIAL_DEPRECIATION_RATE = 0.28

WORKBOOK_ANNUAL_DEPRECIATION_RATE = 0.08

WORKBOOK_DEPRECIATION_YEARS = 10

WORKBOOK_ECONOMIC_LIMIT_CUTOFF = True

def get_years_for_range(start_year, end_year):
    """Return a valid inclusive year list between the chosen range."""
    try:
        start_year = int(start_year)
        end_year = int(end_year)
    except (TypeError, ValueError):
        return YEARS

    if start_year > end_year:
        start_year, end_year = end_year, start_year

    start_year = max(MIN_YEAR, min(start_year, MAX_YEAR))
    end_year = max(MIN_YEAR, min(end_year, MAX_YEAR))
    return list(range(start_year, end_year + 1))

ROW_ORDER = [
    ("input", "Oil_Sales"),
    ("input", "Gas_Sales"),
    ("input", "Condy_Sales"),
    ("calc", "Oil Sales (MMSTB)"),
    ("calc", "Gas Sales (Bscf)"),
    ("calc", "Condy Sales (MMSTB)"),
    ("calc", "Exploration Capex (USD mill)"),
    ("input", "Exploration_Capex"),
    ("calc", "Appraisal Capex (USD mill)"),
    ("input", "Appraisal_Wells"),
    ("calc", "Dev Capex (USD mill)"),
    ("input", "Jacket"),
    ("input", "MachineryandPipelines"),
    ("input", "Drilling"),
    ("calc", "Opex (USD mill)"),
    ("input", "Fixed_Opex"),
    ("input", "Variable_Opex"),
    ("input", "lease"),
    ("input", "Well_PnA"),
    ("calc", "Abex (USD mill)"),
    ("input", "Abex_Cess"),
]

EXCEL_PATH = os.path.join(APP_DIR, "database.xlsx")

NOTIONAL_EXCEL_PATH = os.path.join(
    APP_DIR, "database_notional.xlsx"
)

PRICE_SCENARIOS_PATH = os.path.join(
    APP_DIR, "price_scenarios.json"
)

PROFILE_EXCEL_PATHS = {
    "complete": EXCEL_PATH,
    "notional": NOTIONAL_EXCEL_PATH,
}
_PROFILE_SAVE_LOCK = threading.Lock()

NAMED_DATABASE_PATTERN = re.compile(
    r"^database_[a-z0-9_-]{1,64}_(?:complete|notional)_[0-9a-f]{10}\.xlsx$"
)

PROFILER_DATA = {
    profile_type: {key: {} for key, _ in VARIABLES}
    for profile_type in PROFILE_TYPES
}

def get_profile_type(value):
    """Validate and normalize the selected input profile."""
    profile_type = str(value or "complete").strip().lower()
    if profile_type not in PROFILE_TYPES:
        raise ValueError("Input selection must be Notional or Complete.")
    return profile_type


def get_profile_workbook_path(profile_type="complete"):
    """Resolve the current workbook, including safely named profile files."""
    profile_type = get_profile_type(profile_type)
    legacy_path = PROFILE_EXCEL_PATHS[profile_type]
    directory = os.path.dirname(legacy_path)
    candidates = [legacy_path] if os.path.isfile(legacy_path) else []
    try:
        entries = os.scandir(directory)
    except OSError:
        if candidates:
            return candidates[0]
        raise

    with entries:
        for entry in entries:
            if not entry.is_file() or not NAMED_DATABASE_PATTERN.fullmatch(entry.name):
                continue
            try:
                workbook = load_workbook(entry.path, read_only=True, data_only=True)
                try:
                    if PROJECT_INFO_SHEET not in workbook.sheetnames:
                        continue
                    if str(workbook[PROJECT_INFO_SHEET]["B2"].value or "").strip().lower() == profile_type:
                        candidates.append(entry.path)
                finally:
                    workbook.close()
            except (
                OSError,
                ValueError,
                InvalidFileException,
                zipfile.BadZipFile,
            ) as error:
                raise OSError(
                    f"Could not inspect saved profiler workbook {entry.name}: {error}"
                ) from error

    if not candidates:
        return legacy_path
    return max(candidates, key=os.path.getmtime)


def _named_database_path(project_name, profile_type, directory):
    normalized_name = unicodedata.normalize("NFKD", project_name)
    ascii_name = normalized_name.encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^a-zA-Z0-9_-]+", "_", ascii_name).strip("_-").lower()
    slug = slug[:64].rstrip("_-") or "project"
    digest = hashlib.sha256(
        f"{profile_type}\0{project_name}".encode("utf-8")
    ).hexdigest()[:10]
    filename = f"database_{slug}_{profile_type}_{digest}.xlsx"
    return os.path.join(directory, filename)

def load_price_scenarios():
    """Load the annual price vectors extracted read-only from Price Input.xlsx."""
    if not os.path.exists(PRICE_SCENARIOS_PATH):
        raise FileNotFoundError(
            f"The extracted price data was not found at {PRICE_SCENARIOS_PATH}."
        )
    with open(PRICE_SCENARIOS_PATH, encoding="utf-8") as price_file:
        price_data = json.load(price_file)
    mrp = float(price_data["default_mrp"])
    rows = price_data["rows"]

    scenarios = {name: {} for name in ("base", "low", "high")}
    scenario_indexes = {"base": (1, 2), "low": (3, 4), "high": (5, 6)}
    if not math.isfinite(mrp) or mrp < 0:
        raise ValueError("The workbook's gas-price MRP factor is invalid.")

    for row in rows:
        if len(row) != 7:
            raise ValueError("Price data rows must contain a year and three prices.")
        year = int(row[0])
        for name, (liquid_index, gas_index) in scenario_indexes.items():
            if not MIN_YEAR <= year <= MAX_YEAR:
                continue
            liquid_price = float(row[liquid_index])
            raw_gas_price = float(row[gas_index])
            if not all(math.isfinite(item) for item in (liquid_price, raw_gas_price)):
                raise ValueError(
                    f"The {name} price data for {year} is not numeric."
                )
            scenarios[name][year] = (
                liquid_price,
                raw_gas_price * mrp,
            )

    for name, prices in scenarios.items():
        if set(prices) != set(YEARS):
            raise ValueError(
                f"Price Input.xlsx has incomplete {PRICE_SCENARIO_LABELS[name]} "
                "data for 2020–2060."
            )
    return {"scenarios": scenarios, "default_mrp": mrp}

def resolve_prices(
    years,
    scenario,
    mrp,
    custom_liquid_price,
    custom_gas_price,
    custom_reference_year,
    custom_escalation_rate,
):
    """Resolve year-aligned prices using the workbook scenario/escalation rules."""
    if scenario not in PRICE_SCENARIOS:
        raise ValueError("Price selection must be Base, Low, High, or Custom.")

    source = load_price_scenarios()
    if scenario != "custom":
        price_map = source["scenarios"][scenario]
        return [
            (price_map[year][0], price_map[year][1] / source["default_mrp"] * mrp)
            if source["default_mrp"]
            else (price_map[year][0], 0.0)
            for year in years
        ]

    if not MIN_YEAR <= custom_reference_year <= MAX_YEAR:
        raise ValueError(
            f"Custom price reference year must be between {MIN_YEAR} and {MAX_YEAR}."
        )
    default_liquid_price, default_gas_price = source["scenarios"]["base"][
        custom_reference_year
    ]
    if custom_liquid_price is None:
        custom_liquid_price = default_liquid_price
    if custom_gas_price is None:
        custom_gas_price = default_gas_price
    if not 0 <= custom_escalation_rate <= 1:
        raise ValueError("Custom price escalation must be between 0% and 100%.")
    if custom_liquid_price < 0 or custom_gas_price < 0:
        raise ValueError("Custom prices cannot be negative.")
    return [
        (
            custom_liquid_price
            * (1 + custom_escalation_rate) ** max(0, year - custom_reference_year),
            custom_gas_price
            * (1 + custom_escalation_rate) ** max(0, year - custom_reference_year),
        )
        for year in years
    ]

def load_existing_data(years=None, profile_type="complete"):
    """Read the selected profile workbook so the form can be pre-filled."""
    path = get_profile_workbook_path(profile_type)
    if not os.path.exists(path):
        return {}
    # Data sheet has 2 header rows: leap-check row + year row;
    # a leading "Total" column may also be present.
    df = pd.read_excel(path, sheet_name="Data", index_col=0, header=[0, 1])

    year_columns = {}
    for column in df.columns:
        year_value = column[1] if isinstance(column, tuple) else column
        if str(year_value).strip().isdigit():
            year_columns[int(year_value)] = column
    valid_years = list(year_columns)

    selected_years = years or valid_years
    selected_years = [int(y) for y in selected_years if int(y) in valid_years]

    data = {}
    for key, label in VARIABLES:
        if label not in df.index:
            continue
        data[key] = {}
        row_position = df.index.get_indexer_for([label])[0]
        for y in selected_years:
            if y not in valid_years:
                continue
            val = df.iloc[row_position, df.columns.get_loc(year_columns[y])]
            value = 0 if pd.isna(val) else val
            data[key][int(y)] = value
    return data


def load_project_name(profile_type="complete"):
    """Read the saved project label from the selected profile workbook."""
    path = get_profile_workbook_path(profile_type)
    if not os.path.exists(path):
        return ""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if PROJECT_INFO_SHEET not in workbook.sheetnames:
            return ""
        return str(workbook[PROJECT_INFO_SHEET]["B1"].value or "").strip()
    finally:
        workbook.close()

def leap_indicator(year):
    """Return 0 for leap years, or years since the most recent leap year."""
    year = int(year)
    if calendar.isleap(year):
        return 0

    years_since_leap = 1
    while not calendar.isleap(year - years_since_leap):
        years_since_leap += 1
    return years_since_leap

def days_in_year(year):
    """Return the Gregorian calendar-day count for the year."""
    return 366 if calendar.isleap(int(year)) else 365


def production_volume_factor(year):
    """Convert a daily production rate to that year's million-unit volume."""
    return days_in_year(year) / 1000


PRODUCTION_RATE_KEYS = {"Oil_Sales", "Gas_Sales", "Condy_Sales"}

TOTAL_DECIMAL_PLACES = {"Condy_Sales": 2}

GAS_MCF_PER_BOE = 5.65

DEFAULT_WACC = 0.09

DEFAULT_DISCOUNT_YEAR = 2026

DEFAULT_CUSTOM_REFERENCE_YEAR = 2026

DEFAULT_CUSTOM_ESCALATION_RATE = 0.02

def calculate_project_irr(cash_flows):
    """Calculate an Excel-IRR-style periodic return using a 10% initial guess."""
    if not cash_flows or not (
        any(value < 0 for value in cash_flows)
        and any(value > 0 for value in cash_flows)
    ):
        return None

    rate = 0.1
    for _ in range(100):
        if rate <= -1 or not math.isfinite(rate):
            return None
        base = 1 + rate
        npv = sum(value / base**period for period, value in enumerate(cash_flows))
        derivative = sum(
            -period * value / base ** (period + 1)
            for period, value in enumerate(cash_flows)
            if period
        )
        if abs(npv) < 1e-7:
            return rate
        if derivative == 0:
            return None
        next_rate = rate - npv / derivative
        if not math.isfinite(next_rate) or next_rate <= -1:
            return None
        if abs(next_rate - rate) < 1e-7:
            return next_rate
        rate = next_rate
    return None

def excel_npv(rate, cash_flows):
    """Calculate Excel NPV with the first cash flow occurring one period ahead."""
    return (1 + rate) * sum(
        flow / (1 + rate) ** period
        for period, flow in enumerate(cash_flows, start=1)
    )

def calculate_project_economics(
    years,
    cash_flows,
    wacc=DEFAULT_WACC,
    discount_year=DEFAULT_DISCOUNT_YEAR,
):
    """Calculate project economics at the workbook's discount year."""
    valid_flows = [
        0.0 if value is None or not math.isfinite(float(value)) else float(value)
        for value in cash_flows
    ]
    npv0 = excel_npv(0, valid_flows)
    start_year = years[0] if years else None
    npv_at_wacc = None
    if years and wacc > -1:
        npv_at_wacc = excel_npv(wacc, valid_flows) * (
            1 + wacc
        ) ** (discount_year - start_year)

    economic_limit = (
        max(year for year, flow in zip(years, valid_flows) if flow != 0)
        if any(flow != 0 for flow in valid_flows)
        else start_year
    )
    irr = calculate_project_irr(valid_flows)
    irr_assumed = False
    if irr is None:
        if npv0 > 0 and npv_at_wacc is not None and npv_at_wacc > 0:
            irr = 0.30
            irr_assumed = True
            viability = {
                "status": "strongly-viable",
                "label": "Viable",
                "condition": (
                    f"Green: IRR is N/A and NPV0 and NPV@{wacc:.0%} "
                    "are positive; assume 30%."
                ),
            }
        else:
            viability = {
                "status": "unknown",
                "label": "IRR unavailable",
                "condition": (
                    f"IRR is N/A; assume 30% only when NPV0 and "
                    f"NPV@{wacc:.0%} are both positive."
                ),
            }
    elif irr < 0.11:
        viability = {
            "status": "not-viable",
            "label": "Not viable",
            "condition": "Red: IRR < 11% (not viable).",
        }
    elif irr < 0.20:
        viability = {
            "status": "viable",
            "label": "Viable",
            "condition": "Yellow: 11% ≤ IRR < 20% (viable).",
        }
    else:
        viability = {
            "status": "strongly-viable",
            "label": "Viable",
            "condition": "Green: IRR ≥ 20% (viable).",
        }

    return {
        "npv0": npv0,
        "npv_at_wacc": npv_at_wacc,
        "irr": irr,
        "irr_assumed": irr_assumed,
        "economic_limit": economic_limit,
        "wacc": wacc,
        "discount_year": discount_year,
        "viability": viability,
    }

def calculate_workbook_cash_flows(
    years,
    oil_rates,
    gas_rates,
    condy_rates,
    capex,
    opex,
    abex,
    liquid_capex,
    liquid_opex,
    liquid_abex,
    prices,
    sensitivities,
    revenue_flow_totals=None,
    revenue_flow_rows_output=None,
):
    """Calculate workbook-style annual NCF, including its economic-limit cutoff."""
    series = (
        oil_rates,
        gas_rates,
        condy_rates,
        capex,
        opex,
        abex,
        liquid_capex,
        liquid_opex,
        liquid_abex,
        prices,
    )
    if any(len(values) != len(years) for values in series):
        raise ValueError("Workbook cash-flow inputs must align with the selected years.")

    production_factor = sensitivities["production"]
    capex_factor = sensitivities["capex"]
    opex_factor = sensitivities["opex"]
    abex_factor = sensitivities["abex"]

    production_years = [
        year
        for year, oil, gas, condy in zip(years, oil_rates, gas_rates, condy_rates)
        if oil > 0 or gas > 0 or condy > 0
    ]
    first_production_year = min(production_years) if production_years else None

    scaled_capex = [amount * capex_factor for amount in capex]
    capex_vintages = {}
    if first_production_year is not None:
        preproduction_capex = sum(
            amount
            for year, amount in zip(years, scaled_capex)
            if year < first_production_year
        )
        if preproduction_capex:
            capex_vintages[first_production_year] = preproduction_capex
        for year, amount in zip(years, scaled_capex):
            if year >= first_production_year:
                capex_vintages[year] = capex_vintages.get(year, 0.0) + amount

    qpe_carry_forward = 0.0
    adjusted_loss_carry_forward = 0.0
    cumulative_limit_balance = 0.0
    cumulative_limit_balances = []
    cash_flows = []
    revenue_flow_rows = []
    for index, year in enumerate(years):
        volume_factor = production_volume_factor(year)
        oil_volume = oil_rates[index] * production_factor * volume_factor
        condy_volume = condy_rates[index] * production_factor * volume_factor
        gas_volume = gas_rates[index] * production_factor * volume_factor
        liquid_revenue = (oil_volume + condy_volume) * prices[index][0]
        gas_revenue = gas_volume * prices[index][1]
        revenue = liquid_revenue + gas_revenue
        royalty = WORKBOOK_ROYALTY_RATE * revenue

        annual_capex = scaled_capex[index]
        annual_opex = opex[index] * opex_factor
        annual_abex = abex[index] * abex_factor
        cumulative_limit_balance += revenue - annual_opex
        cumulative_limit_balances.append(cumulative_limit_balance)
        additional_liquid_cost = (
            liquid_capex[index] * capex_factor
            + liquid_opex[index] * opex_factor
            + liquid_abex[index] * abex_factor
        )
        export_duty_cost = (
            annual_capex + annual_opex + annual_abex + additional_liquid_cost
        )
        export_duty = WORKBOOK_EXPORT_DUTY_RATE * max(
            0.0,
            liquid_revenue * (1 - WORKBOOK_ROYALTY_RATE) - export_duty_cost,
        )

        depreciation = 0.0
        if first_production_year is not None and year >= first_production_year:
            for vintage_year, vintage_amount in capex_vintages.items():
                age = year - vintage_year
                if age == 0:
                    depreciation += (
                        vintage_amount * WORKBOOK_INITIAL_DEPRECIATION_RATE
                    )
                elif 0 < age < WORKBOOK_DEPRECIATION_YEARS:
                    depreciation += (
                        vintage_amount * WORKBOOK_ANNUAL_DEPRECIATION_RATE
                    )

        net_operating_income = (
            revenue - royalty - annual_opex - annual_abex - export_duty
        )
        qpe_used = min(
            max(0.0, net_operating_income),
            depreciation + qpe_carry_forward,
        )
        qpe_carry_forward += depreciation - qpe_used
        statutory_income = max(0.0, net_operating_income - qpe_used)
        adjusted_loss = min(
            statutory_income,
            max(0.0, -net_operating_income) + adjusted_loss_carry_forward,
        )
        adjusted_loss_carry_forward += (
            max(0.0, -net_operating_income) - adjusted_loss
        )
        taxable_income = statutory_income - adjusted_loss
        tax_income = (
            revenue
            - royalty
            - depreciation
            - annual_opex
            - annual_abex
            - export_duty
        )
        pita = (
            taxable_income * WORKBOOK_PITA_RATE
            if tax_income > 0
            else 0.0
        )
        project_cost = annual_capex + annual_opex + annual_abex
        project_ncf = (
            revenue
            - royalty
            - project_cost
            - export_duty
            - pita
        )
        cash_flows.append(project_ncf)
        revenue_flow_rows.append({
            "Gross Revenue": revenue,
            "Project Cost": project_cost,
            "Royalty": royalty,
            "Export Duty": export_duty,
            "PITA": pita,
            "Project NCF": project_ncf,
        })

    if WORKBOOK_ECONOMIC_LIMIT_CUTOFF and years:
        economic_limit_index = cumulative_limit_balances.index(
            max(cumulative_limit_balances)
        )
        economic_limit_year = years[economic_limit_index]
        cash_flows = [
            cash_flow if year <= economic_limit_year else 0.0
            for year, cash_flow in zip(years, cash_flows)
        ]
        for index, year in enumerate(years):
            if year > economic_limit_year:
                revenue_flow_rows[index] = {
                    key: 0.0 for key in revenue_flow_rows[index]
                }

    if revenue_flow_totals is not None:
        revenue_flow_totals.clear()
        for key in revenue_flow_rows[0] if revenue_flow_rows else ():
            revenue_flow_totals[key] = sum(row[key] for row in revenue_flow_rows)
    if revenue_flow_rows_output is not None:
        revenue_flow_rows_output.clear()
        revenue_flow_rows_output.extend(revenue_flow_rows)
    return cash_flows

def canonical_label(label):
    """Normalize label text so raw input labels and display labels match."""
    if label is None:
        return ""
    text = str(label)
    text = text.replace("(Mstb/d)", "")
    text = text.replace("(MMscf/d)", "")
    text = text.replace("(USD mill)", "")
    text = text.replace("(MMSTB)", "")
    text = text.replace("(Bscf)", "")
    text = text.replace("(MUSD)", "")
    text = text.replace("&", "and")
    text = text.replace("P&A", "PnA")
    return " ".join(text.split()).strip()

def calculate_metrics(df):
    """
    Run the profiler's derived calculations on top of the raw input grid.

    `df` has one row per variable label (e.g. "Oil Sales") and one
    column per year. Returns a DataFrame in the same shape (rows =
    calculated field names, columns = years) so it can be written to
    its own sheet in the Excel file.
    """
    numeric_df = df.apply(pd.to_numeric, errors="coerce")
    years = df.columns
    lookup = {canonical_label(label): label for label in numeric_df.index}

    volume_factors = pd.Series(
        {year: production_volume_factor(year) for year in years}
    )
    oil_sales_mmstb = (
        numeric_df.loc[lookup[canonical_label("Oil Sales")]] * volume_factors
    )
    gas_sales_bscf = (
        numeric_df.loc[lookup[canonical_label("Gas Sales")]] * volume_factors
    )
    condy_sales_mmstb = (
        numeric_df.loc[lookup[canonical_label("Condy Sales")]] * volume_factors
    )

    exploration_capex = numeric_df.loc[lookup[canonical_label("Exploration Capex")]]
    appraisal_capex = numeric_df.loc[lookup[canonical_label("Appraisal Wells")]]
    dev_capex = (
        numeric_df.loc[lookup[canonical_label("Jacket")]]
        + numeric_df.loc[lookup[canonical_label("Machinery & Pipelines")]]
        + numeric_df.loc[lookup[canonical_label("Drilling")]]
    )
    opex = (
        numeric_df.loc[lookup[canonical_label("Fixed Opex")]]
        + numeric_df.loc[lookup[canonical_label("Variable Opex")]]
        + numeric_df.loc[lookup[canonical_label("Lease")]]
        + numeric_df.loc[lookup[canonical_label("Well P&A")]]
    )
    abex = numeric_df.loc[lookup[canonical_label("Abex Cess")]]

    results = pd.DataFrame({
        "Oil Sales (MMSTB)": oil_sales_mmstb,
        "Gas Sales (Bscf)": gas_sales_bscf,
        "Condy Sales (MMSTB)": condy_sales_mmstb,
        "Exploration Capex (USD mill)": exploration_capex,
        "Appraisal Capex (USD mill)": appraisal_capex,
        "Dev Capex (USD mill)": dev_capex,
        "Opex (USD mill)": opex,
        "Abex (USD mill)": abex,
    }).T
    results.index.name = "Calculated Field"
    return results

def build_raw_input_dataframe(existing, years):
    raw_data = {}
    for key, label in VARIABLES:
        raw_data[label] = [existing.get(key, {}).get(year, "") for year in years]
    raw_df = pd.DataFrame(raw_data, index=years).T
    raw_df.index.name = "Variable"
    return raw_df

def get_display_calculations(existing, years=None):
    """Display the calculated metrics for the given years and saved inputs."""
    if not existing:
        return {}
    selected_years = years or YEARS
    return results_to_display(
        calculate_metrics(build_raw_input_dataframe(existing, selected_years)),
        years=selected_years,
    )

def build_rows():
    """Resolve ROW_ORDER into full row dicts the template can loop over."""
    var_labels = dict(VARIABLES)
    rows = []
    for kind, ref in ROW_ORDER:
        label = var_labels[ref] if kind == "input" else ref
        display_label, _, unit = label.rpartition(" (")
        if not display_label or not unit.endswith(")"):
            raise ValueError(f"Profiler row label has no unit: {label}")
        unit = unit[:-1]
        rows.append({
            "type": kind,
            "key": ref if kind == "input" else None,
            "label": label,
            "display_label": {
                "Appraisal Wells": "Appraisal Well",
            }.get(display_label, display_label),
            "unit": {
                "MMSTB": "MMstb",
                "USD mill": "USD mil",
            }.get(unit, unit),
        })
    return rows

def results_to_display(results, years=None):
    """Turn a calculate_metrics() DataFrame into {label: {year: value_or_blank}}."""
    selected_years = years or YEARS
    display = {}
    for label in results.index:
        display[label] = {}
        for year in selected_years:
            val = results.loc[label, year]
            if pd.isna(val):
                display[label][year] = ""
            else:
                numeric_val = float(val)
                display[label][year] = "" if abs(numeric_val) < 1e-9 else round(numeric_val, 1)
    return display

def format_total_value(total, decimal_places=1):
    """Return a display-friendly total with the requested precision."""
    if pd.isna(total):
        return "-"
    total = float(total)
    if abs(total) < 1e-9:
        return "0"
    return f"{total:.{decimal_places}f}"

def calculate_row_total(key, years, values, calculated=False):
    """Total annual values, converting raw production rates to annual volumes."""
    total = 0.0
    has_values = False
    for year, value in zip(years, values):
        if value in (None, ""):
            continue
        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            continue
        if pd.isna(numeric_value) or (calculated and abs(numeric_value) < 1e-9):
            continue
        if key in PRODUCTION_RATE_KEYS and not calculated:
            numeric_value *= production_volume_factor(year)
        total += numeric_value
        has_values = True
    return total if has_values else None

def build_total_display(rows, years, existing, calc_display):
    """Return a mapping of row key/label to a total across the active year range."""
    totals = {}
    calculated_totals = (
        calculate_metrics(build_raw_input_dataframe(existing, years))
        if existing else None
    )
    for row in rows:
        key = row["key"] if row["type"] == "input" else row["label"]
        if row["type"] == "input":
            values = [existing.get(row["key"], {}).get(year, "") for year in years]
            total_key = key
            calculated = False
        else:
            values = (
                calculated_totals.loc[row["label"]].tolist()
                if calculated_totals is not None
                else [calc_display.get(row["label"], {}).get(year, "") for year in years]
            )
            total_key = key
            calculated = True

        total = calculate_row_total(total_key, years, values, calculated=calculated)
        totals[key] = (
            format_total_value(total, TOTAL_DECIMAL_PLACES.get(total_key, 1))
            if total is not None
            else "-"
        )
    return totals

def save_form_data(form, years=None, profile_type="complete"):
    """Serialize profile saves so workbook discovery and renames cannot race."""
    with _PROFILE_SAVE_LOCK:
        return _save_form_data(form, years=years, profile_type=profile_type)


def _save_form_data(form, years=None, profile_type="complete"):
    """
    Save the current profile data and return the raw input DataFrame.
    """
    profile_type = get_profile_type(profile_type)
    project_name = str(form.get("project_name", "") or "").strip()
    if not project_name:
        raise ValueError("Project name is required.")
    if len(project_name) > 120:
        raise ValueError("Project name must be 120 characters or fewer.")
    source_path = get_profile_workbook_path(profile_type)
    output_path = _named_database_path(
        project_name, profile_type, os.path.dirname(source_path)
    )
    if os.path.exists(output_path) and os.path.abspath(output_path) != os.path.abspath(source_path):
        raise ValueError(
            "A database workbook with the requested safe filename already exists. "
            "Rename or move that file before saving this project."
        )
    selected_years = years or YEARS
    selected_year_set = set(selected_years)
    saved_data = load_existing_data(years=YEARS, profile_type=profile_type)

    data = {}
    for key, label in VARIABLES:
        row = []
        for year in YEARS:
            if year in selected_year_set:
                raw_value = form.get(f"{key}_{year}", "")
            else:
                raw_value = saved_data.get(key, {}).get(year, "")
            if raw_value is None or str(raw_value).strip() == "":
                row.append(0)
            else:
                try:
                    row.append(round(float(raw_value), 1))
                except (TypeError, ValueError):
                    row.append(0.0)
        data[label] = row

    df = pd.DataFrame(data, index=YEARS).T
    df.index.name = "Variable"

    results = calculate_metrics(df)
    economic_sheets = economic_data(df, results)

    # Keep the raw inputs and the calculated outputs in the same main data
    # sheet so Excel shows them alongside each other, while still retaining
    # the separate Calculations sheet for a clean read-only view.
    combined_df = pd.concat([df, results], axis=0)
    combined_df = combined_df.apply(pd.to_numeric, errors="coerce")
    combined_df.index.name = "Variable"

    input_keys = {label: key for key, label in VARIABLES}
    total_values = []
    for label in combined_df.index:
        if label in input_keys:
            values = [df.loc[label, year] for year in YEARS]
            total_key = input_keys[label]
            calculated = False
        else:
            values = results.loc[label].tolist()
            total_key = label
            calculated = True

        row_total = calculate_row_total(
            total_key, YEARS, values, calculated=calculated
        )
        total_values.append(
            round(row_total, TOTAL_DECIMAL_PLACES.get(total_key, 1))
            if row_total is not None
            else "-"
        )

    df_for_excel = combined_df.round(1).copy()
    df_for_excel.insert(0, "Total", total_values)

    excel_years = ["Total"] + list(df_for_excel.columns[1:])
    excel_leap = ["Total"] + [leap_indicator(y) for y in df_for_excel.columns[1:]]
    df_for_excel.columns = pd.MultiIndex.from_arrays(
        [excel_leap, excel_years],
        names=["Leap Check (0=leap year; otherwise years since)", "Variable"],
    )

    temporary_descriptor, temporary_path = tempfile.mkstemp(
        prefix=".icon-profiler-", suffix=".xlsx", dir=os.path.dirname(output_path)
    )
    os.close(temporary_descriptor)
    try:
        with pd.ExcelWriter(temporary_path, engine="openpyxl") as writer:
            df_for_excel.to_excel(writer, sheet_name="Data")
            results.round(1).to_excel(writer, sheet_name="Calculations")
            for sheet_name, sheet_df in economic_sheets.items():
                sheet_df.round(1).to_excel(writer, sheet_name=sheet_name, index=False)
            project_info = writer.book.create_sheet(PROJECT_INFO_SHEET)
            project_info["A1"] = "Project Name"
            project_info["B1"] = project_name
            project_info["A2"] = "Profile Type"
            project_info["B2"] = profile_type
            writer.book.properties.title = project_name

            data_start_rows = {"Data": 4, "Calculations": 2}
            for worksheet in writer.book.worksheets:
                start_row = data_start_rows.get(worksheet.title, 2)
                for row in worksheet.iter_rows(min_row=start_row):
                    for cell in row:
                        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                            cell.number_format = (
                                "#,##0" if worksheet.title == "Revenue Flow" else "0.0"
                            )
                if worksheet.title == "Revenue Flow":
                    for column, width in {
                        "A": 22,
                        "B": 16,
                        "C": 12,
                        "D": 15,
                        "E": 10,
                    }.items():
                        worksheet.column_dimensions[column].width = width
                elif worksheet.title == "Project Net Cash Flow":
                    worksheet.column_dimensions["A"].width = 12
                    worksheet.column_dimensions["B"].width = 24
                    worksheet.column_dimensions["C"].width = 30
                    for cell in worksheet["A"][1:]:
                        cell.number_format = "0"
            data_sheet = writer.book["Data"]
            for label, key in input_keys.items():
                if key == "Condy_Sales":
                    for row_label in (label, "Condy Sales (MMSTB)"):
                        row = 4 + combined_df.index.get_loc(row_label)
                        data_sheet.cell(row=row, column=2).number_format = "0.00"

        os.replace(temporary_path, output_path)
        if os.path.abspath(source_path) != os.path.abspath(output_path):
            try:
                os.remove(source_path)
            except FileNotFoundError:
                pass
            except OSError as error:
                try:
                    os.remove(output_path)
                except OSError as rollback_error:
                    raise OSError(
                        "The project-named workbook was created, but the previous "
                        "workbook could not be removed and the new file could not "
                        "be rolled back."
                    ) from rollback_error
                raise OSError(
                    "Could not rename the previous profiler workbook. Close it in "
                    "Excel or wait for file syncing to finish, then try again."
                ) from error
    finally:
        if os.path.exists(temporary_path):
            os.remove(temporary_path)

    return df, results

def economic_data(raw_data, calculated_data):
    """Build category-specific, multi-year economic data sheets."""
    years = list(raw_data.columns)

    def values(label):
        return pd.to_numeric(raw_data.loc[label], errors="coerce").fillna(0)

    def calculated(label):
        return pd.to_numeric(calculated_data.loc[label], errors="coerce").fillna(0)

    def category(columns):
        sheet = pd.DataFrame(
            {year: [series.get(year, 0) for _, series in columns] for year in years},
            index=[column_name for column_name, _ in columns],
        )
        sheet.insert(0, "Category", sheet.index)
        sheet["Total"] = sheet[years].sum(axis=1)
        return sheet.reset_index(drop=True)

    days = pd.Series({year: days_in_year(year) for year in years})
    oil_rate = values("Oil Sales (Mstb/d)")
    gas_rate = values("Gas Sales (MMscf/d)")
    condy_rate = values("Condy Sales (Mstb/d)")

    fixed_variable_opex = values("Fixed Opex (USD mill)") + values("Variable Opex (USD mill)")
    lease_opex = values("Lease (USD mill)")
    well_pna_opex = values("Well P&A (USD mill)")

    years_for_cash_flow = list(raw_data.columns)
    price_source = load_price_scenarios()
    prices = resolve_prices(
        years_for_cash_flow,
        "base",
        price_source["default_mrp"],
        None,
        None,
        DEFAULT_CUSTOM_REFERENCE_YEAR,
        DEFAULT_CUSTOM_ESCALATION_RATE,
    )
    revenue_flow_totals = {}
    project_cash_flows = calculate_workbook_cash_flows(
        years_for_cash_flow,
        values("Oil Sales (Mstb/d)").tolist(),
        values("Gas Sales (MMscf/d)").tolist(),
        values("Condy Sales (Mstb/d)").tolist(),
        (
            values("Drilling (USD mill)")
            + values("Jacket (USD mill)")
            + values("Machinery & Pipelines (USD mill)")
            + values("Appraisal Wells (USD mill)")
            + values("Exploration Capex (USD mill)")
        ).tolist(),
        (
            values("Fixed Opex (USD mill)")
            + values("Variable Opex (USD mill)")
            + values("Lease (USD mill)")
            + values("Well P&A (USD mill)")
        ).tolist(),
        values("Abex Cess (USD mill)").tolist(),
        values("Liquid Capex for Export Duty (USD mill)").tolist(),
        values("Liquid Opex for Export Duty (USD mill)").tolist(),
        values("Liquid Abex for Export Duty (USD mill)").tolist(),
        prices,
        {"production": 1.0, "capex": 1.0, "opex": 1.0, "abex": 1.0},
        revenue_flow_totals=revenue_flow_totals,
    )
    project_ncf_sheet = pd.DataFrame({
        "Year": years_for_cash_flow,
        "Project NCF (USD mil)": project_cash_flows,
        "Cum Project NCF (USD mil)": pd.Series(project_cash_flows).cumsum(),
    })

    revenue_flow = pd.DataFrame([
        {
            "Mil (RM or USD)": "Gross Revenue",
            "Project Cost": revenue_flow_totals["Gross Revenue"],
            "Royalty": None,
            "Export Duty": None,
            "PITA": None,
        },
        {
            "Mil (RM or USD)": "Cost",
            "Project Cost": revenue_flow_totals["Project Cost"],
            "Royalty": revenue_flow_totals["Royalty"],
            "Export Duty": revenue_flow_totals["Export Duty"],
            "PITA": revenue_flow_totals["PITA"],
        },
        {
            "Mil (RM or USD)": "Project NCF",
            "Project Cost": revenue_flow_totals["Project NCF"],
            "Royalty": None,
            "Export Duty": None,
            "PITA": None,
        },
    ])

    return {
        "Liquid Production": category([
            ("Oil Rate (Mstb/d)", oil_rate),
            ("Oil Volume (Mstb)", oil_rate * days),
            ("Condensate Rate (Mstb/d)", condy_rate),
            ("Condensate Volume (Mstb)", condy_rate * days),
        ]),
        "Gas Production": category([
            ("Gas Rate (MMscf/d)", gas_rate),
            ("Gas Volume (MMscf)", gas_rate * days),
        ]),
        "Opex": category([
            ("Fixed & Variable Opex (USD mill)", fixed_variable_opex),
            ("Lease (USD mill)", lease_opex),
            ("Oil Well Abandonment Cost (USD mill)", well_pna_opex),
            (
                "Total Opex (USD mill)",
                fixed_variable_opex + lease_opex + well_pna_opex,
            ),
        ]),
        "Capex": category([
            ("Oil Exploration Well (USD mill)", values("Exploration Capex (USD mill)")),
            ("Oil Appraisal Well (USD mill)", values("Appraisal Wells (USD mill)")),
            ("Oil Platform / Jacket / Substructure (USD mill)", values("Jacket (USD mill)")),
            ("Oil Plant Machinery (USD mill)", values("Machinery & Pipelines (USD mill)")),
            ("Oil Development Well (USD mill)", values("Drilling (USD mill)")),
            (
                "Total Oil Capex (USD mill)",
                values("Exploration Capex (USD mill)")
                + values("Appraisal Wells (USD mill)")
                + values("Jacket (USD mill)")
                + values("Machinery & Pipelines (USD mill)")
                + values("Drilling (USD mill)"),
            ),
        ]),
        "Abex": category([
            ("Abex Cess (USD mill)", values("Abex Cess (USD mill)")),
            ("Total Abex (USD mill)", calculated("Abex (USD mill)")),
        ]),
        "Revenue Flow": revenue_flow,
        "Project Net Cash Flow": project_ncf_sheet,
    }

def econimic_data(raw_data=None, calculated_data=None):
    """Backward-compatible spelling for callers using the original name."""
    if raw_data is None or calculated_data is None:
        path = get_profile_workbook_path("complete")
        if not os.path.exists(path):
            return {}
        return pd.ExcelFile(path).sheet_names
    return economic_data(raw_data, calculated_data)

def build_dashboard_data(
    years,
    wacc=DEFAULT_WACC,
    profile_type="complete",
    price_scenario="base",
    mrp=0.45,
    custom_liquid_price=None,
    custom_gas_price=None,
    custom_reference_year=DEFAULT_CUSTOM_REFERENCE_YEAR,
    custom_escalation_rate=DEFAULT_CUSTOM_ESCALATION_RATE,
    sensitivities=None,
    cash_flow_mode="workbook",
):
    """Build dashboard series and workbook-style project economics."""
    profile_type = get_profile_type(profile_type)
    if cash_flow_mode not in {"workbook", "manual"}:
        raise ValueError("Cash-flow basis must be Workbook or Manual.")
    if sensitivities is None:
        sensitivities = {
            "production": 1.0,
            "capex": 1.0,
            "opex": 1.0,
            "abex": 1.0,
        }

    fiscal_years = YEARS
    selected_years = set(years)
    existing = load_existing_data(years=fiscal_years, profile_type=profile_type)
    for key, _ in VARIABLES:
        existing.setdefault(key, {})
        existing[key].update(
            {
                year: value
                for year, value in PROFILER_DATA[profile_type].get(key, {}).items()
                if year in fiscal_years
            }
        )

    fiscal_prices = resolve_prices(
        fiscal_years,
        price_scenario,
        mrp,
        custom_liquid_price,
        custom_gas_price,
        custom_reference_year,
        custom_escalation_rate,
    )
    price_by_year = dict(zip(fiscal_years, fiscal_prices))
    resolved_prices = [price_by_year[year] for year in years]
    reference_price = resolve_prices(
        [custom_reference_year],
        price_scenario,
        mrp,
        custom_liquid_price,
        custom_gas_price,
        custom_reference_year,
        custom_escalation_rate,
    )[0]

    def value(key, year):
        raw_value = existing.get(key, {}).get(year, 0)
        if raw_value is None or (
            isinstance(raw_value, str) and not raw_value.strip()
        ):
            return 0.0
        try:
            number = float(raw_value)
        except (TypeError, ValueError) as error:
            raise ValueError(
                f"{key} for {year} must be a finite number."
            ) from error
        if not math.isfinite(number):
            raise ValueError(f"{key} for {year} must be a finite number.")
        return number

    charts = {
        "oil": {"label": "Oil Production (kbbl/d)", "series": [{"name": "Oil", "values": []}]},
        "condy": {"label": "Condy Production (kbbl/d)", "series": [{"name": "Condy", "values": []}]},
        "gas": {"label": "Gas Production (mmscf/d)", "series": [{"name": "Gas", "values": []}]},
        "capex": {"label": "Capex (USD mil)", "series": [
            {"name": "Dev - Wells", "values": []},
            {"name": "Dev Facilities", "values": []},
            {"name": "Appraisal", "values": []},
            {"name": "Exploration", "values": []},
        ]},
        "opex": {"label": "Opex (USD mil)", "series": [
            {"name": "Well P&A", "values": []},
            {"name": "Lease", "values": []},
            {"name": "Fixed & Variable", "values": []},
        ]},
        "abex": {"label": "Abex (USD mil)", "series": [{"name": "Abex - Facilities", "values": []}]},
    }
    totals = {"oil_mmstb": 0.0, "gas_bscf": 0.0, "condy_mmstb": 0.0,
              "capex": 0.0, "opex": 0.0, "abex": 0.0,
              "exploration": 0.0}
    annual_inputs = {
        key: []
        for key in (
            "oil_rates",
            "gas_rates",
            "condy_rates",
            "capex",
            "opex",
            "abex",
            "liquid_capex",
            "liquid_opex",
            "liquid_abex",
            "manual_cash_flows",
        )
    }

    for year in fiscal_years:
        oil = value("Oil_Sales", year)
        gas = value("Gas_Sales", year)
        condy = value("Condy_Sales", year)
        drilling = value("Drilling", year)
        facilities = value("Jacket", year) + value("MachineryandPipelines", year)
        appraisal = value("Appraisal_Wells", year)
        exploration = value("Exploration_Capex", year)
        well_pna = value("Well_PnA", year)
        lease = value("lease", year)
        fixed_variable = value("Fixed_Opex", year) + value("Variable_Opex", year)
        abex = value("Abex_Cess", year)
        annual_inputs["oil_rates"].append(oil)
        annual_inputs["gas_rates"].append(gas)
        annual_inputs["condy_rates"].append(condy)
        annual_inputs["capex"].append(drilling + facilities + appraisal + exploration)
        annual_inputs["opex"].append(well_pna + lease + fixed_variable)
        annual_inputs["abex"].append(abex)
        annual_inputs["liquid_capex"].append(value("Liquid_Capex", year))
        annual_inputs["liquid_opex"].append(value("Liquid_Opex", year))
        annual_inputs["liquid_abex"].append(value("Liquid_Abex", year))
        annual_inputs["manual_cash_flows"].append(value("Project_NCF", year))

        if year not in selected_years:
            continue
        charts["oil"]["series"][0]["values"].append(oil)
        charts["condy"]["series"][0]["values"].append(condy)
        charts["gas"]["series"][0]["values"].append(gas)
        for series, amount in zip(charts["capex"]["series"],
                                  (drilling, facilities, appraisal, exploration)):
            series["values"].append(amount)
        for series, amount in zip(charts["opex"]["series"], (well_pna, lease, fixed_variable)):
            series["values"].append(amount)
        charts["abex"]["series"][0]["values"].append(abex)

        volume_factor = production_volume_factor(year)
        totals["oil_mmstb"] += oil * volume_factor
        totals["gas_bscf"] += gas * volume_factor
        totals["condy_mmstb"] += condy * volume_factor
        totals["capex"] += drilling + facilities + appraisal + exploration
        totals["exploration"] += exploration
        totals["opex"] += well_pna + lease + fixed_variable
        totals["abex"] += abex

    fiscal_revenue_rows = []
    if cash_flow_mode == "manual":
        manual_cash_flow_by_year = dict(
            zip(fiscal_years, annual_inputs["manual_cash_flows"])
        )
        cash_flow_by_year = manual_cash_flow_by_year
        project_cash_flows = [cash_flow_by_year[year] for year in years]
    else:
        fiscal_cash_flows = calculate_workbook_cash_flows(
            fiscal_years,
            annual_inputs["oil_rates"],
            annual_inputs["gas_rates"],
            annual_inputs["condy_rates"],
            annual_inputs["capex"],
            annual_inputs["opex"],
            annual_inputs["abex"],
            annual_inputs["liquid_capex"],
            annual_inputs["liquid_opex"],
            annual_inputs["liquid_abex"],
            fiscal_prices,
            sensitivities,
            revenue_flow_rows_output=fiscal_revenue_rows,
        )
        cash_flow_by_year = dict(zip(fiscal_years, fiscal_cash_flows))
        project_cash_flows = [cash_flow_by_year[year] for year in years]

    cash_flow_chart_years = list(years)
    cash_flow_chart_values = [
        cash_flow_by_year[year] for year in cash_flow_chart_years
    ]
    cumulative_cash_flow_chart_values = []
    cumulative_cash_flow = 0.0
    for cash_flow in cash_flow_chart_values:
        cumulative_cash_flow += cash_flow
        cumulative_cash_flow_chart_values.append(cumulative_cash_flow)

    revenue_row_by_year = dict(zip(fiscal_years, fiscal_revenue_rows))
    revenue_split = {
        key: sum(revenue_row_by_year[year][key] for year in years)
        for key in (
            "Gross Revenue",
            "Project Cost",
            "Royalty",
            "Export Duty",
            "PITA",
            "Project NCF",
        )
    } if cash_flow_mode == "workbook" else {}

    sensitivity_analysis = {}
    if cash_flow_mode == "workbook":
        sensitivity_keys = {
            "production": "production",
            "capex": "capex",
            "opex": "opex",
        }
        for label, sensitivity_key in sensitivity_keys.items():
            sensitivity_analysis[label] = {}
            for case, multiplier in (("minus", 0.6), ("base", 1.0), ("plus", 1.4)):
                case_sensitivities = dict(sensitivities)
                case_sensitivities[sensitivity_key] *= multiplier
                case_cash_flows = calculate_workbook_cash_flows(
                    fiscal_years,
                    annual_inputs["oil_rates"],
                    annual_inputs["gas_rates"],
                    annual_inputs["condy_rates"],
                    annual_inputs["capex"],
                    annual_inputs["opex"],
                    annual_inputs["abex"],
                    annual_inputs["liquid_capex"],
                    annual_inputs["liquid_opex"],
                    annual_inputs["liquid_abex"],
                    fiscal_prices,
                    case_sensitivities,
                )
                case_cash_flow_by_year = dict(zip(fiscal_years, case_cash_flows))
                selected_cash_flows = [
                    case_cash_flow_by_year[year] for year in years
                ]
                case_economics = calculate_project_economics(
                    years, selected_cash_flows, wacc=wacc
                )
                sensitivity_analysis[label][case] = {
                    "irr": case_economics["irr"],
                    "npv0": case_economics["npv0"],
                }

    total_mmboe = (
        totals["oil_mmstb"]
        + totals["condy_mmstb"]
        + totals["gas_bscf"] / GAS_MCF_PER_BOE
    )

    def ratio(numerator, denominator):
        return numerator / denominator if denominator else None

    summary = {
        "production": {
            "oil": totals["oil_mmstb"],
            "gas": totals["gas_bscf"],
            "condy": totals["condy_mmstb"],
            "total": total_mmboe,
        },
        "costs": {
            key: totals[key]
            for key in ("capex", "opex", "abex", "exploration")
        },
        "unit_costs": {
            "ufc": ratio(totals["exploration"], total_mmboe),
            "udc": ratio(totals["capex"], total_mmboe),
            "upc": ratio(totals["opex"], total_mmboe),
            "uac": ratio(totals["abex"], total_mmboe),
            "utc": ratio(
                totals["exploration"]
                + totals["capex"]
                + totals["opex"]
                + totals["abex"],
                total_mmboe,
            ),
            "utc_gas": ratio(
                totals["capex"]
                + totals["opex"]
                + totals["abex"],
                totals["gas_bscf"],
            ),
        },
    }
    return {
        "years": years,
        "charts": charts,
        "conversions": {"gas_mcf_per_boe": GAS_MCF_PER_BOE},
        "prices": {
            "scenario": price_scenario,
            "label": PRICE_SCENARIO_LABELS[price_scenario],
            "years": years,
            "liquid_usd_bbl": [price[0] for price in resolved_prices],
            "gas_usd_mmbtu": [price[1] for price in resolved_prices],
            "mrp": mrp,
            "default_mrp": load_price_scenarios()["default_mrp"],
            "reference_year": custom_reference_year,
            "gas_price_at_reference": reference_price[1],
            "cash_flow_mode": cash_flow_mode,
        },
        "project_cash_flows": project_cash_flows,
        "project_cash_flow_chart": {
            "years": cash_flow_chart_years,
            "project_ncf": cash_flow_chart_values,
            "cumulative_project_ncf": cumulative_cash_flow_chart_values,
        },
        "revenue_split": revenue_split,
        "sensitivity_analysis": sensitivity_analysis,
        "project_economics": calculate_project_economics(
            years, project_cash_flows, wacc=wacc
        ),
        "summary": summary,
    }

def update_profiler_snapshot(form, years, profile_type="complete"):
    """Copy the current profiler form values into the in-memory dashboard source."""
    profile_type = get_profile_type(profile_type)
    for key, _ in VARIABLES:
        for year in years:
            field_name = f"{key}_{year}"
            if field_name not in form:
                continue
            raw_value = form.get(field_name, "")
            try:
                PROFILER_DATA[profile_type][key][year] = (
                    round(float(raw_value), 1) if str(raw_value).strip() else 0.0
                )
            except (TypeError, ValueError):
                PROFILER_DATA[profile_type][key][year] = 0.0

def production(oil_rate,gas_rate,condy_rate):
    """Return input production rates and their combined oil-equivalent rate."""
    oil_production = float(oil_rate)
    gas_production = float(gas_rate)
    condy_production = float(condy_rate)
    total_production = (
        oil_production + gas_production / GAS_MCF_PER_BOE + condy_production
    )
    return {
        "oil production": oil_production,
        "gas production": gas_production,
        "condy production": condy_production,
        "total production": total_production,
    }

def costs(total_oil_capex, total_opex, total_abex): 
    capex = total_oil_capex
    opex = total_opex
    abex = total_abex
    return capex,opex,abex

def Summary(
    total_production=0,
    gas_production=0,
    capex=0,
    opex=0,
    abex=0,
    exploration_capex=0,
):
    """Return summary economic ratios with zero-safe division."""

    def safe_divide(numerator, denominator):
        return 0.0 if denominator == 0 else numerator / denominator

    UFC = safe_divide(exploration_capex, total_production)
    UDC = safe_divide(capex, total_production)
    UPC = safe_divide(opex, total_production)
    UAC = safe_divide(abex, total_production)
    UTC = UFC + UDC + UPC + UAC
    UTC_Gas = safe_divide(capex + opex + abex, gas_production)

    return {
        "UFC": UFC,
        "UDC": UDC,
        "UPC": UPC,
        "UAC": UAC,
        "UTC": UTC,
        "UTC_Gas": UTC_Gas,
    }
