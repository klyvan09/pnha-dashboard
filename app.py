
# ============================================================
# PNHA CHE DASHBOARD
# Revised layout, compact sidebar, dropdown text fixes
# ============================================================
#
# Workbook:
# H:\My Drive\PNHA CHE Dashboard\CHE_Database.xlsx
#
# Run:
#   python app.py
#
# Install dependencies if needed:
#   pip install dash pandas openpyxl plotly
# ============================================================

from pathlib import Path
from math import ceil, floor, log10
import re

import pandas as pd
import plotly.graph_objects as go

from dash import (
    Dash,
    Input,
    Output,
    State,
    ALL,
    ctx,
    dcc,
    html,
    no_update,
)


# ============================================================
# 1. CONFIGURATION
# ============================================================

EXCEL_PATH = Path(__file__).resolve().with_name("CHE_Database.xlsx")

YEARS = list(range(2014, 2026))

SHEET_NAMES = [
    "FS", "FSRI", "FA", "HF", "FP", "HP", "HC",
    "AGESEX", "DIS", "REG", "INC", "HK",
]

SPECIAL_SHEET_NAMES = ["PHC", "PERCAPITA"]

CHE_CODES = [
    "FS", "FSRI", "FA", "HF", "FP", "HP", "HC",
    "AGESEX", "DIS", "REG", "INC",
]

HK_CODE = "HK"

CLASSIFICATION_LABELS = {
    "FS": "Revenue of Health Financing Scheme",
    "FSRI": "Institutional Unit Providing Revenues to Financing Scheme",
    "FA": "Financing Agent",
    "HF": "Health Care Financing Scheme",
    "FP": "Factors of Health Care Provision",
    "HP": "Health Care Provider",
    "HC": "Health Care Function",
    "AGESEX": "Age",
    "DIS": "Diseases",
    "REG": "Regions",
    "INC": "Income Groups",
    "HK": "Health Capital Formation",
}

STACK_CLASSIFICATION_LABELS = {
    "FS": "Revenue of Health Financing Scheme",
    "FSRI": "Institutional Unit Providing Revenues to Financing Scheme",
    "FA": "Financing Agent",
    "HF": "Health Care Financing Scheme",
    "FP": "Factors of Health Care Provision",
    "HP": "Health Care Provider",
    "HC": "Health Care Function",
    "AGESEX": "Age",
    "DIS": "Diseases",
    "REG": "Regions",
    "INC": "Income Groups",
}

CHE_CLASSIFICATION_DEFINITION_TITLES = {
    "FS": "By Revenue of Health Financing Scheme",
    "FSRI": "By Institutional Unit Providing Revenues to Financing Scheme",
    "FA": "By Financing Agent",
    "HF": "By Health Care Financing Scheme",
    "FP": "By Factors of Health Care Provision",
    "HP": "By Health Care Provider",
    "HC": "By Health Care Function",
    "AGESEX": "By Age",
    "DIS": "By Diseases",
    "REG": "By Regions",
    "INC": "By Income Groups",
}

CHE_CLASSIFICATION_DEFINITIONS = {
    "FS": "Presents the different sources of funds used to finance health care, showing how financing schemes obtain the resources needed to support health-related expenditures.",
    "FSRI": "Presents the different institutions or entities that provide financial resources to health financing schemes, illustrating the contributions of various sectors to health care financing.",
    "FA": "Presents the entities responsible for managing, administering, or channeling financial resources for health care, showing how funds are handled within the health financing system.",
    "HF": "Presents the different mechanisms and arrangements through which health care is financed, providing an overview of how financial resources are organized to support access to health services.",
    "FP": "Presents the different resources and inputs used in delivering health care services, illustrating how expenditures are allocated to the resources needed for service delivery.",
    "HP": "Presents the different institutions, establishments, and professionals that deliver health care services, showing how health expenditure is distributed among providers.",
    "HC": "Presents the different types of health care services and activities provided, illustrating how health expenditure is allocated according to the purpose of health care.",
    "AGESEX": "Presents the distribution of health expenditure according to the age groups of individuals receiving or benefiting from health care services, showing how health resources are allocated across different stages of life.",
    "DIS": "Presents the distribution of health expenditure according to the diseases and health conditions of individuals receiving health care, highlighting how resources are allocated to address different health needs.",
    "REG": "Presents the geographical distribution of health expenditure across regions, providing an overview of how health spending is distributed throughout the country.",
    "INC": "Presents the distribution of health expenditure according to the income groups of individuals receiving or benefiting from health care services, showing how health resources are distributed across different economic groups.",
}

CHE_ANALYSIS_FOCUS = {
    "FS": (
        "Transfers from government domestic revenue (allocated to health purposes)",
        "Transfers from government domestic revenue (allocated to health purposes)",
    ),
    "FSRI": ("Government", "Government"),
    "FA": ("Households", "Households"),
    "HF": ("Household out-of-pocket payment", "Household out-of-pocket payment"),
    "FP": ("Consumption of fixed capital", "Consumption of fixed capital"),
    "HP": ("Providers of preventive care", "Providers of preventive care"),
    "HC": ("Curative care", "Curative care"),
    "AGESEX": ("65 and over", "Age 65 and over"),
    "DIS": ("Noncommunicable diseases", "Noncommunicable diseases"),
    "REG": ("NCR", "NCR"),
    "INC": ("Fifth Quintile", "Fifth Quintile"),
}

CHE_ANALYSIS_ALIASES = {
    "HF": ["Household out-of-pocket payment", "household out of pocket payment", "oop"],
    "FP": ["Consumption of fixed capital"],
    "HP": ["Providers of preventive care"],
    "AGESEX": ["65 and over", "65+"],
    "DIS": ["Noncommunicable diseases", "Non-communicable diseases"],
}

CHE_ANALYSIS_SECONDARY = {
    "FS": ["Other domestic revenues"],
    "FSRI": ["Households"],
    "FA": ["Central government"],
    "HF": ["Government schemes"],
    "FP": ["Materials and services used"],
    "HP": ["Hospitals"],
    "HC": ["Preventive care"],
}

TOP_REGIONS_TO_SHOW = 7
TOP_CATEGORIES_TO_SHOW = 7


# ============================================================
# 2. LOAD AND NORMALIZE WORKBOOK
# ============================================================

if not EXCEL_PATH.exists():
    raise FileNotFoundError(
        f"Excel workbook not found:\n{EXCEL_PATH}"
    )


def clean_column_name(column):
    return str(column).strip()


def find_year_columns(df):
    year_columns = {}

    for column in df.columns:
        text = str(column).strip()
        match = re.fullmatch(r"(20\d{2})(?:\.0)?", text)

        if match:
            year = int(match.group(1))
            if year in YEARS:
                year_columns[year] = column

        elif isinstance(column, (int, float)):
            year = int(column)
            if year in YEARS:
                year_columns[year] = column

    return year_columns


def find_column(df, possible_names):
    lookup = {
        str(column).strip().lower(): column
        for column in df.columns
    }

    for name in possible_names:
        if name.lower() in lookup:
            return lookup[name.lower()]

    return None


def normalize_dataframe(df, sheet_name):
    df = df.copy()
    df.columns = [clean_column_name(c) for c in df.columns]

    component_column = find_column(
        df,
        ["Component", "Classification", "Category", "Name"],
    )

    subcomponent_column = find_column(
        df,
        [
            "Sub-component",
            "Sub Component",
            "Subcomponent",
            "Sub-category",
            "Subcategory",
        ],
    )

    if component_column is None:
        year_columns_before = find_year_columns(df)
        non_year_columns = [
            c for c in df.columns
            if c not in year_columns_before.values()
        ]

        if not non_year_columns:
            raise ValueError(
                f"Sheet '{sheet_name}' has no classification column."
            )

        component_column = non_year_columns[0]

    rename_map = {component_column: "Component"}

    if subcomponent_column is not None:
        rename_map[subcomponent_column] = "Sub-component"

    df = df.rename(columns=rename_map)

    year_columns = find_year_columns(df)
    df = df.rename(
        columns={
            original: str(year)
            for year, original in year_columns.items()
        }
    )

    for year in YEARS:
        year_column = str(year)

        if year_column not in df.columns:
            df[year_column] = 0.0

        cleaned = (
            df[year_column]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("₱", "", regex=False)
            .str.strip()
            .replace(
                {
                    "": None,
                    "-": None,
                    "nan": None,
                    "None": None,
                    "N/A": None,
                }
            )
        )

        df[year_column] = pd.to_numeric(
            cleaned,
            errors="coerce",
        ).fillna(0.0)

    df["Component"] = (
        df["Component"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    if "Sub-component" not in df.columns:
        df["Sub-component"] = ""

    df["Sub-component"] = (
        df["Sub-component"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df = df[
        (df["Component"] != "")
        | (df["Sub-component"] != "")
    ].copy()

    df["Sheet"] = sheet_name

    return df.reset_index(drop=True)


RAW_SHEETS = pd.read_excel(
    EXCEL_PATH,
    sheet_name=None,
    engine="openpyxl",
)

DATA = {}

for sheet_name in SHEET_NAMES + SPECIAL_SHEET_NAMES:
    if sheet_name in RAW_SHEETS:
        DATA[sheet_name] = normalize_dataframe(
            RAW_SHEETS[sheet_name],
            sheet_name,
        )
    else:
        print(
            f"WARNING: Worksheet '{sheet_name}' was not found."
        )

if "FS" not in DATA:
    raise ValueError(
        "The FS worksheet is required to calculate CHE."
    )

if "HK" not in DATA:
    print(
        "WARNING: HK worksheet missing. "
        "Health Capital Formation will be treated as zero."
    )


# ============================================================
# 3. TOTAL CALCULATIONS
# ============================================================

TOTAL_LABELS = {
    "total",
    "grand total",
    "overall total",
    "total current health expenditure",
    "total health expenditure",
    "total health capital formation",
}


def get_total_from_sheet(sheet_name, year):
    """
    Use an explicit total row if available.
    Otherwise, sum the non-total rows.

    Verify the fallback against official totals if the worksheet
    contains both parent aggregates and their child rows.
    """

    if sheet_name not in DATA:
        return 0.0

    df = DATA[sheet_name]
    year_column = str(year)

    category_text = (
        df["Component"].astype(str).str.strip().str.lower()
    )
    subcategory_text = (
        df["Sub-component"].astype(str).str.strip().str.lower()
    )

    explicit_total = (
        category_text.isin(TOTAL_LABELS)
        | subcategory_text.isin(TOTAL_LABELS)
    )

    if explicit_total.any():
        return float(
            df.loc[explicit_total, year_column].iloc[0]
        )

    eligible_rows = (
        ~explicit_total
        & (df["Component"].str.strip() != "")
        & (df["Component"].str.lower() != "nan")
    )

    return float(df.loc[eligible_rows, year_column].sum())


CHE_SERIES = pd.Series(
    {
        year: get_total_from_sheet("FS", year)
        for year in YEARS
    },
    dtype=float,
)

HK_SERIES = pd.Series(
    {
        year: get_total_from_sheet("HK", year)
        for year in YEARS
    },
    dtype=float,
)

THE_SERIES = CHE_SERIES.add(
    HK_SERIES,
    fill_value=0.0,
)


def get_component_series(sheet_name, component_name):
    sheet_df = DATA.get(sheet_name)
    if sheet_df is None:
        return pd.Series({year: 0.0 for year in YEARS}, dtype=float)

    matching_rows = sheet_df[
        sheet_df["Component"].str.strip().str.lower()
        == component_name.strip().lower()
    ]
    if matching_rows.empty:
        return pd.Series({year: 0.0 for year in YEARS}, dtype=float)

    return pd.Series(
        {
            year: float(matching_rows[str(year)].sum())
            for year in YEARS
        },
        dtype=float,
    )


PHC_SERIES = get_component_series(
    "PHC",
    "Primary Health Care Expenditure",
)
PHC_CHE_SERIES = get_component_series(
    "PHC",
    "Current Health Expenditure",
)
PERCAPITA_SERIES = get_component_series(
    "PERCAPITA",
    "Current Health Expenditure Per Capita",
)

METRIC_LABELS = {
    "THE": "Total Health Expenditure",
    "CHE": "Current Health Expenditure",
    "PERCAPITA": "Current Health Expenditure per Capita",
    "HK": "Health Capital Formation",
    "PHC": "Primary Health Care",
}

METRIC_SERIES = {
    "THE": THE_SERIES,
    "CHE": CHE_SERIES,
    "PERCAPITA": PERCAPITA_SERIES,
    "HK": HK_SERIES,
    "PHC": PHC_SERIES,
}


# ============================================================
# 4. CATEGORY DATA AND GROUPING
# ============================================================

def get_raw_category_data(sheet_name, year):
    if sheet_name not in DATA:
        return pd.DataFrame(columns=["Category", "Amount"])

    df = DATA[sheet_name].copy()
    year_column = str(year)

    df["Category"] = df["Sub-component"].where(
        df["Sub-component"].str.strip() != "",
        df["Component"],
    )

    df["Category"] = df["Category"].astype(str).str.strip()

    category_lower = df["Category"].str.lower()

    df = df[
        (df["Category"] != "")
        & (category_lower != "nan")
        & (~category_lower.isin(TOTAL_LABELS))
    ].copy()

    df["Amount"] = pd.to_numeric(
        df[year_column],
        errors="coerce",
    ).fillna(0.0)

    df = df[df["Amount"] > 0].copy()

    if df.empty:
        return pd.DataFrame(columns=["Category", "Amount"])

    return (
        df.groupby("Category", as_index=False)["Amount"]
        .sum()
        .sort_values("Amount", ascending=False)
        .reset_index(drop=True)
    )


def get_named_category_amount(sheet_name, year, aliases):
    """Return a category total using normalized exact or contained aliases."""
    category_df = get_raw_category_data(sheet_name, year)
    normalized_aliases = [
        re.sub(r"[^a-z0-9]", "", alias.casefold())
        for alias in aliases
    ]
    matches = category_df["Category"].map(
        lambda name: any(
            alias == re.sub(r"[^a-z0-9]", "", str(name).casefold())
            or alias in re.sub(r"[^a-z0-9]", "", str(name).casefold())
            for alias in normalized_aliases
        )
    ) if not category_df.empty else pd.Series(dtype=bool)
    return float(category_df.loc[matches, "Amount"].sum()) if not category_df.empty else 0.0


def group_other_regions(category_df):
    """Keep the largest seven regions; combine the rest."""

    if category_df.empty:
        return category_df

    category_df = category_df.copy()

    existing_other = category_df[
        category_df["Category"].str.strip().str.lower()
        == "other regions"
    ]

    remaining = category_df[
        category_df["Category"].str.strip().str.lower()
        != "other regions"
    ].copy()

    remaining = remaining.sort_values(
        "Amount",
        ascending=False,
    ).reset_index(drop=True)

    top_regions = remaining.head(TOP_REGIONS_TO_SHOW)
    other_regions = remaining.iloc[TOP_REGIONS_TO_SHOW:]

    other_amount = float(other_regions["Amount"].sum())

    if not existing_other.empty:
        other_amount += float(existing_other["Amount"].sum())

    result = top_regions.copy()

    if other_amount > 0:
        result = pd.concat(
            [
                result,
                pd.DataFrame(
                    [{
                        "Category": "Other Regions",
                        "Amount": other_amount,
                    }]
                ),
            ],
            ignore_index=True,
        )

    return result.sort_values(
        "Amount",
        ascending=False,
    ).reset_index(drop=True)


def group_other_categories(category_df):
    """Keep the largest seven categories; combine the rest."""

    if category_df.empty:
        return category_df

    category_df = category_df.copy()

    existing_other = category_df[
        category_df["Category"].str.strip().str.lower()
        == "other categories"
    ]

    remaining = category_df[
        category_df["Category"].str.strip().str.lower()
        != "other categories"
    ].copy()

    remaining = remaining.sort_values(
        "Amount",
        ascending=False,
    ).reset_index(drop=True)

    top_categories = remaining.head(TOP_CATEGORIES_TO_SHOW)
    other_categories = remaining.iloc[TOP_CATEGORIES_TO_SHOW:]

    other_amount = float(other_categories["Amount"].sum())

    if not existing_other.empty:
        other_amount += float(existing_other["Amount"].sum())

    result = top_categories.copy()

    if other_amount > 0:
        result = pd.concat(
            [
                result,
                pd.DataFrame(
                    [{
                        "Category": "Other categories",
                        "Amount": other_amount,
                    }]
                ),
            ],
            ignore_index=True,
        )

    return result.sort_values(
        "Amount",
        ascending=False,
    ).reset_index(drop=True)


def get_category_data(sheet_name, year):
    category_df = get_raw_category_data(sheet_name, year)

    if category_df.empty:
        return category_df

    if sheet_name == "REG":
        return group_other_regions(category_df)

    return group_other_categories(category_df)


def is_focus_category(classification_code, category):
    aliases = CHE_ANALYSIS_ALIASES.get(
        classification_code,
        [CHE_ANALYSIS_FOCUS[classification_code][0]],
    )
    normalized_category = re.sub(r"[^a-z0-9]", "", str(category).lower())
    normalized_aliases = [
        re.sub(r"[^a-z0-9]", "", alias.lower())
        for alias in aliases
    ]
    return any(
        alias == normalized_category or alias in normalized_category
        for alias in normalized_aliases
    )


def is_secondary_focus_category(classification_code, category):
    aliases = CHE_ANALYSIS_SECONDARY.get(classification_code, [])
    normalized_category = re.sub(r"[^a-z0-9]", "", str(category).lower())
    return any(
        alias == normalized_category or alias in normalized_category
        for alias in (re.sub(r"[^a-z0-9]", "", item.lower()) for item in aliases)
    )


def make_che_analysis_text(classification_code, start_year, end_year):
    """Build a data-driven classification insight for the displayed year range."""
    analysis_years = [year for year in YEARS if start_year <= year <= end_year]
    if not analysis_years:
        analysis_years = [end_year]
    first_year, last_year = analysis_years[0], analysis_years[-1]
    focus_label = CHE_ANALYSIS_FOCUS[classification_code][1]

    def series_for(category_code, category):
        values = {}
        for year in analysis_years:
            if category_code == "FA" and category in {"General government", "Households"}:
                df = DATA["FA"]
                mask = df["Component"].str.strip().str.casefold() == category.casefold()
                amount = float(pd.to_numeric(df.loc[mask, str(year)], errors="coerce").fillna(0).sum())
            else:
                df = get_raw_category_data(category_code, year)
                aliases = CHE_ANALYSIS_ALIASES.get(category_code, [category]) if category == CHE_ANALYSIS_FOCUS[category_code][0] else [category]
                normalized_aliases = [re.sub(r"[^a-z0-9]", "", alias.casefold()) for alias in aliases]
                matches = df["Category"].map(
                    lambda name: any(
                        alias == re.sub(r"[^a-z0-9]", "", str(name).casefold())
                        or alias in re.sub(r"[^a-z0-9]", "", str(name).casefold())
                        for alias in normalized_aliases
                    )
                ) if not df.empty else pd.Series(dtype=bool)
                amount = float(df.loc[matches, "Amount"].sum()) if not df.empty else 0.0
            values[year] = amount
        return pd.Series(values, dtype=float)

    def amount_share(category_code, category, year):
        amount = float(series_for(category_code, category).get(year, 0.0))
        total = float(CHE_SERIES.get(year, 0.0))
        return amount, (amount / total * 100 if total else 0.0)

    def pct_change(start, end):
        return ((end / start) - 1) * 100 if start else 0.0

    def strong(text):
        return html.Strong(text)

    span = f"{first_year}–{last_year}"

    if classification_code == "FS":
        transfer = series_for("FS", CHE_ANALYSIS_FOCUS["FS"][0])
        other_revenues = series_for("FS", "Other domestic revenues")
        transfer_shares = pd.Series({y: amount_share("FS", CHE_ANALYSIS_FOCUS["FS"][0], y)[1] for y in analysis_years})
        other_shares = pd.Series({y: amount_share("FS", "Other domestic revenues", y)[1] for y in analysis_years})
        transfer_growth = (transfer[last_year] / transfer[first_year] - 1) * 100 if transfer[first_year] else 0
        other_growth = (other_revenues[last_year] / other_revenues[first_year] - 1) * 100 if other_revenues[first_year] else 0
        transfer_trend = "rose" if transfer[last_year] >= transfer[first_year] else "fell"
        other_trend = "rose" if other_revenues[last_year] >= other_revenues[first_year] else "fell"
        transfer_share_trend = "rose" if transfer_shares[last_year] >= transfer_shares[first_year] else "fell"
        other_share_trend = "fell" if other_shares[last_year] <= other_shares[first_year] else "rose"
        transfer_gap_start = other_shares[first_year] - transfer_shares[first_year]
        transfer_gap_end = other_shares[last_year] - transfer_shares[last_year]
        return html.Span([
            strong("Government transfers and Other Domestic Revenues"),
            f": government transfers {transfer_trend} from {format_amount(transfer[first_year])} in {first_year} "
            f"to {format_amount(transfer[last_year])} in {last_year} ({transfer_growth:+.1f}%), compared with "
            f"Other Domestic Revenues, which {other_trend} from {format_amount(other_revenues[first_year])} "
            f"to {format_amount(other_revenues[last_year])} ({other_growth:+.1f}%). The transfer’s share "
            f"of total FS revenues {transfer_share_trend} from {transfer_shares[first_year]:.1f}% to "
            f"{transfer_shares[last_year]:.1f}%, while Other Domestic Revenues’ share {other_share_trend} "
            f"from {other_shares[first_year]:.1f}% to {other_shares[last_year]:.1f}%; the gap narrowed "
            f"from {transfer_gap_start:.1f} to {transfer_gap_end:.1f} percentage points.",
        ])
    if classification_code in {"FSRI", "FA"}:
        government = series_for(classification_code, "Government" if classification_code == "FSRI" else "General government")
        household = series_for(classification_code, "Households")
        gov_shares = pd.Series({y: government[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        hh_shares = pd.Series({y: household[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        gov_up_years = sum(gov_shares[y] > gov_shares[analysis_years[i - 1]] for i, y in enumerate(analysis_years) if i > 0)
        hh_down_years = sum(hh_shares[y] < hh_shares[analysis_years[i - 1]] for i, y in enumerate(analysis_years) if i > 0)
        intervals = max(len(analysis_years) - 1, 0)
        if classification_code == "FSRI":
            return html.Span([
                strong("Government and household funding"),
                f": government funding rose from {format_amount(government[first_year])} in {first_year} "
                f"to {format_amount(government[last_year])} in {last_year}, outpacing household "
                f"funding, which rose from {format_amount(household[first_year])} to "
                f"{format_amount(household[last_year])}. Government’s share of CHE increased "
                f"from {gov_shares[first_year]:.1f}% to {gov_shares[last_year]:.1f}%, while "
                f"the household share shrank from {hh_shares[first_year]:.1f}% to "
                f"{hh_shares[last_year]:.1f}%.",
            ])
        if classification_code == "FA":
            central_government = series_for("FA", "Central government")
            household_shares = pd.Series({y: household[y] / CHE_SERIES[y] * 100 for y in analysis_years})
            central_shares = pd.Series({y: central_government[y] / CHE_SERIES[y] * 100 for y in analysis_years})
            household_growth = (household[last_year] / household[first_year] - 1) * 100 if household[first_year] else 0
            central_growth = (central_government[last_year] / central_government[first_year] - 1) * 100 if central_government[first_year] else 0
            household_trend = "rose" if household[last_year] >= household[first_year] else "fell"
            central_trend = "rose" if central_government[last_year] >= central_government[first_year] else "fell"
            household_share_trend = "fell" if household_shares[last_year] < household_shares[first_year] else "rose"
            central_share_trend = "rose" if central_shares[last_year] > central_shares[first_year] else "fell"
            central_gap_start = round(household_shares[first_year], 1) - round(central_shares[first_year], 1)
            central_gap_end = round(household_shares[last_year], 1) - round(central_shares[last_year], 1)
            relative_growth = "faster than" if central_growth > household_growth else "slower than" if central_growth < household_growth else "at the same rate as"
            gap_trend = "narrowed" if central_gap_end < central_gap_start else "widened" if central_gap_end > central_gap_start else "held steady"
            return html.Span([
                strong("Households and Central Government as financing agents"),
                f": household financing {household_trend} from {format_amount(household[first_year])} in {first_year} "
                f"to {format_amount(household[last_year])} in {last_year} ({household_growth:+.1f}%), while "
                f"Central Government {central_trend} from {format_amount(central_government[first_year])} to "
                f"{format_amount(central_government[last_year])} ({central_growth:+.1f}%). Central Government "
                f"funding grew {relative_growth} household financing. Household share of CHE {household_share_trend} "
                f"from {household_shares[first_year]:.1f}% to {household_shares[last_year]:.1f}%, while Central "
                f"Government share {central_share_trend} from {central_shares[first_year]:.1f}% to "
                f"{central_shares[last_year]:.1f}%; the share gap {gap_trend} from {central_gap_start:.1f} "
                f"to {central_gap_end:.1f} percentage points.",
            ])
        label = "Government" if classification_code == "FSRI" else "General government"
        classification_name = "institutional units providing revenues" if classification_code == "FSRI" else "financing agents"
        return html.Span([
            strong(f"Government and household shares among {classification_name}"),
            f": government’s share of CHE rose overall from {gov_shares[first_year]:.1f}% "
            f"in {first_year} to {gov_shares[last_year]:.1f}% in {last_year}, increasing in "
            f"{gov_up_years} of {intervals} year-to-year changes. Household share moved "
            f"from {hh_shares[first_year]:.1f}% to {hh_shares[last_year]:.1f}% and declined "
            f"in {hh_down_years} of {intervals} changes.",
        ])

    if classification_code == "HF":
        government = series_for("HF", "Government schemes")
        household = series_for("HF", CHE_ANALYSIS_FOCUS["HF"][0])
        gov_shares = pd.Series({y: government[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        hh_shares = pd.Series({y: household[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        gov_growth = (government[last_year] / government[first_year] - 1) * 100 if government[first_year] else 0
        hh_growth = (household[last_year] / household[first_year] - 1) * 100 if household[first_year] else 0
        gov_share_trend = "rose" if gov_shares[last_year] >= gov_shares[first_year] else "fell"
        hh_share_trend = "fell" if hh_shares[last_year] <= hh_shares[first_year] else "rose"
        relative_growth = "faster" if gov_growth > hh_growth else "slower" if gov_growth < hh_growth else "at the same rate as"
        return html.Span([
            strong("Government schemes and household out-of-pocket payments"),
            f": government scheme funding changed from {format_amount(government[first_year])} in {first_year} "
            f"to {format_amount(government[last_year])} in {last_year} ({gov_growth:+.1f}%), growing {relative_growth} "
            f"than household out-of-pocket funding ({hh_growth:+.1f}%). Government schemes’ share of CHE "
            f"{gov_share_trend} from {gov_shares[first_year]:.1f}% to {gov_shares[last_year]:.1f}%, while "
            f"household out-of-pocket share {hh_share_trend} from {hh_shares[first_year]:.1f}% "
            f"to {hh_shares[last_year]:.1f}%.",
        ])
    if classification_code == "FP":
        fixed_capital = series_for("FP", CHE_ANALYSIS_FOCUS["FP"][0])
        materials = series_for("FP", "Materials and services used")
        fixed_share = pd.Series({y: fixed_capital[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        material_share = pd.Series({y: materials[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        fixed_share_trend = "fell" if fixed_share[last_year] < fixed_share[first_year] else "rose"
        material_share_trend = "rose" if material_share[last_year] > material_share[first_year] else "fell"
        fixed_level_trend = "rose" if fixed_capital[last_year] > fixed_capital[first_year] else "fell"
        materials_level_trend = "rose" if materials[last_year] > materials[first_year] else "fell"
        return html.Span([
            strong("Consumption of fixed capital and materials and services used"),
            f": consumption of fixed capital’s share of CHE {fixed_share_trend} from "
            f"{fixed_share[first_year]:.1f}% in {first_year} to {fixed_share[last_year]:.1f}% "
            f"in {last_year}, even as its level {fixed_level_trend} from {format_amount(fixed_capital[first_year])} "
            f"to {format_amount(fixed_capital[last_year])}. The share for materials and services used "
            f"{material_share_trend} from {material_share[first_year]:.1f}% to "
            f"{material_share[last_year]:.1f}%, while its level {materials_level_trend} from "
            f"{format_amount(materials[first_year])} to {format_amount(materials[last_year])}. "
            f"Together, the share trends show CHE shifting toward materials and services relative to "
            f"fixed-capital consumption; a declining share does not mean CFC spending fell in level terms.",
        ])
    if classification_code == "HP":
        preventive = series_for("HP", CHE_ANALYSIS_FOCUS["HP"][0])
        hospitals = series_for("HP", "Hospitals")
        preventive_share = pd.Series({y: preventive[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        hospital_share = pd.Series({y: hospitals[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        preventive_share_trend = "fell" if preventive_share[last_year] < preventive_share[first_year] else "rose" if preventive_share[last_year] > preventive_share[first_year] else "held steady"
        hospital_share_trend = "fell" if hospital_share[last_year] < hospital_share[first_year] else "rose" if hospital_share[last_year] > hospital_share[first_year] else "held steady"
        preventive_growth = pct_change(preventive[first_year], preventive[last_year])
        hospital_growth = pct_change(hospitals[first_year], hospitals[last_year])
        return html.Span([
            strong("Providers of preventive care and hospitals"),
            f": preventive-care providers’ spending changed from {format_amount(preventive[first_year])} "
            f"in {first_year} to {format_amount(preventive[last_year])} in {last_year} "
            f"({preventive_growth:+.1f}%), while its share of CHE {preventive_share_trend} "
            f"from {preventive_share[first_year]:.1f}% to {preventive_share[last_year]:.1f}%. "
            f"Hospital spending changed from {format_amount(hospitals[first_year])} to "
            f"{format_amount(hospitals[last_year])} ({hospital_growth:+.1f}%), and its share "
            f"{hospital_share_trend} from {hospital_share[first_year]:.1f}% to "
            f"{hospital_share[last_year]:.1f}%. This compares the preventive-care provider share "
            f"with hospitals over the selected period.",
        ])
    if classification_code == "HC":
        curative = series_for("HC", "Curative care")
        preventive = series_for("HC", "Preventive care")
        curative_shares = pd.Series({y: curative[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        preventive_shares = pd.Series({y: preventive[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        curative_trend = "fell" if curative_shares[last_year] < curative_shares[first_year] else "rose" if curative_shares[last_year] > curative_shares[first_year] else "held steady"
        preventive_trend = "fell" if preventive_shares[last_year] < preventive_shares[first_year] else "rose" if preventive_shares[last_year] > preventive_shares[first_year] else "held steady"
        curative_change = curative_shares[last_year] - curative_shares[first_year]
        preventive_change = preventive_shares[last_year] - preventive_shares[first_year]
        return html.Span([
            strong("Curative and preventive care shares of CHE"),
            f": curative care’s share {curative_trend} from {curative_shares[first_year]:.1f}% "
            f"in {first_year} to {curative_shares[last_year]:.1f}% in {last_year} "
            f"({curative_change:+.1f} percentage points). Preventive care’s share {preventive_trend} "
            f"from {preventive_shares[first_year]:.1f}% to {preventive_shares[last_year]:.1f}% "
            f"({preventive_change:+.1f} percentage points). Curative care remained the larger share "
            f"throughout the selected period, ranging from {curative_shares.min():.1f}% to "
            f"{curative_shares.max():.1f}%, while preventive care ranged from "
            f"{preventive_shares.min():.1f}% to {preventive_shares.max():.1f}%.",
        ])
    if classification_code == "AGESEX":
        focus = series_for("AGESEX", "65 and over")
        ranks = {}
        for year in analysis_years:
            categories = get_raw_category_data("AGESEX", year)
            ranked = categories.sort_values("Amount", ascending=False).reset_index(drop=True)
            match = ranked[ranked["Category"].str.casefold() == "65 and over"]
            ranks[year] = int(match.index[0] + 1) if not match.empty else len(ranked) + 1
        top_years = sum(rank == 1 for rank in ranks.values())
        return html.Span([
            strong("Age 65 and over"),
            f" had the largest age-group share in {top_years} of {len(analysis_years)} "
            f"years in the selected period; its share was "
            f"{focus[first_year] / CHE_SERIES[first_year] * 100:.1f}% in {first_year} and "
            f"{focus[last_year] / CHE_SERIES[last_year] * 100:.1f}% in {last_year}.",
        ])

    if classification_code == "REG":
        ncr_top = 0
        calabarzon_top_two = 0
        ncr_shares = {}
        calabarzon_shares = {}
        for year in analysis_years:
            categories = get_raw_category_data("REG", year)
            categories = categories[~categories["Category"].str.casefold().isin(["nationwide"])]
            names = categories["Category"].tolist()
            ncr_rank = next((i + 1 for i, name in enumerate(names) if "ncr" in name.casefold()), None)
            calabarzon_rank = next((i + 1 for i, name in enumerate(names) if "calabarzon" in name.casefold()), None)
            if ncr_rank == 1:
                ncr_top += 1
            if calabarzon_rank is not None and calabarzon_rank <= 2:
                calabarzon_top_two += 1
            ncr_shares[year] = amount_share("REG", "NCR", year)[1]
            calabarzon_shares[year] = amount_share("REG", "CALABARZON", year)[1]
        return html.Span([
            strong("NCR and CALABARZON"),
            f" were among the leading regions: NCR ranked first in {ncr_top} of "
            f"{len(analysis_years)} years, while CALABARZON ranked in the top two in "
            f"{calabarzon_top_two}. Their CHE shares changed from "
            f"{ncr_shares[first_year]:.1f}% to {ncr_shares[last_year]:.1f}% for NCR and "
            f"{calabarzon_shares[first_year]:.1f}% to {calabarzon_shares[last_year]:.1f}% "
            f"for CALABARZON. (Nationwide is excluded from regional rankings.)",
        ])

    if classification_code == "DIS":
        amounts = series_for("DIS", CHE_ANALYSIS_FOCUS["DIS"][0])
        shares = pd.Series({y: amounts[y] / CHE_SERIES[y] * 100 for y in analysis_years})
        return html.Span([
            strong("Noncommunicable diseases"),
            f" increased its share of CHE overall from {shares[first_year]:.1f}% in "
            f"{first_year} to {shares[last_year]:.1f}% in {last_year}, despite "
            f"year-to-year fluctuations.",
        ])

    if classification_code == "INC":
        focus = series_for("INC", "Fifth Quintile")
        top_years = 0
        for year in analysis_years:
            categories = get_raw_category_data("INC", year)
            if not categories.empty and categories.iloc[0]["Category"].casefold() == "fifth quintile":
                top_years += 1
        return html.Span([
            strong("Fifth income quintile"),
            f" had the largest share in all {top_years} years of the selected period; its share was "
            f"{focus[first_year] / CHE_SERIES[first_year] * 100:.1f}% in {first_year} and "
            f"{focus[last_year] / CHE_SERIES[last_year] * 100:.1f}% in {last_year}.",
        ])

    return html.Span([strong(focus_label), f" across {span}."])


def get_selected_series(code):
    if code == HK_CODE:
        return HK_SERIES.copy()

    return CHE_SERIES.copy()


def format_amount(value):
    magnitude = abs(value)
    if magnitude >= 1_000_000:
        scaled_value = value / 1_000_000
        unit = "trillion"
    elif magnitude >= 1_000:
        scaled_value = value / 1_000
        unit = "billion"
    else:
        scaled_value = value
        unit = "million"

    return f"PhP {scaled_value:,.2f} {unit}"


def format_kpi_amount(value):
    return format_amount(value)


def scale_amounts_for_axis(values):
    magnitude = float(values.abs().max()) if len(values) else 0.0
    if magnitude >= 1_000_000:
        return values / 1_000_000, "trillion"
    if magnitude >= 1_000:
        return values / 1_000, "billion"
    return values, "million"


def get_category_color_map(category_df):
    palette = [
        "#4C9F38",
        "#78B866",
        "#2F7D32",
        "#A0C989",
        "#5F9E50",
        "#B9D8A9",
        "#3F8541",
        "#8FBD7C",
        "#6AA653",
        "#C8E0BC",
    ]
    return {
        category: palette[index % len(palette)]
        for index, category in enumerate(category_df["Category"].astype(str))
    }


# ============================================================
# 5. CHART STYLE HELPERS
# ============================================================

def make_empty_figure(message):
    fig = go.Figure()

    fig.add_annotation(
        text=message,
        x=0.5,
        y=0.5,
        xref="paper",
        yref="paper",
        showarrow=False,
        font=dict(size=13, color="#f3f6fa"),
    )

    fig.update_layout(
        template="plotly_white",
        margin=dict(l=20, r=20, t=35, b=20),
        paper_bgcolor="#2a4568",
        plot_bgcolor="#2a4568",
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
    )

    return fig


def apply_chart_style(fig, title=None):
    fig.update_layout(
        template="plotly_white",
        title=dict(
            text=title or "",
            x=0.02,
            xanchor="left",
            font=dict(size=15, color="#f3f6fa"),
        ),
        margin=dict(l=35, r=25, t=55, b=45),
        font=dict(
            family="Arial, sans-serif",
            size=11,
            color="#e6edf7",
        ),
        paper_bgcolor="#2a4568",
        plot_bgcolor="#2a4568",
        hoverlabel=dict(
            bgcolor="#f8fafc",
            font_size=12,
            font=dict(color="#172033"),
        ),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.15,
            xanchor="left",
            x=0,
        ),
    )

    fig.update_xaxes(
        showgrid=False,
        zeroline=False,
        linecolor="#60738e",
        tickfont=dict(color="#e6edf7"),
        title_font=dict(color="#e6edf7"),
    )

    fig.update_yaxes(
        showgrid=True,
        gridcolor="#3b4f6c",
        zeroline=False,
        tickfont=dict(color="#e6edf7"),
        title_font=dict(color="#e6edf7"),
    )

    return fig


def add_chart_heading(fig, title, subtitle, top_margin=62):
    """Add a consistent title and subtitle above a chart."""
    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b><br><sup>{subtitle}</sup>",
            x=0.01,
            xanchor="left",
            y=0.92,
            yanchor="top",
            font=dict(size=20, color="#f3f6fa"),
        ),
        margin=dict(t=top_margin),
    )
    return fig


def set_time_series_y_axis(fig, values, view):
    """Use four evenly spaced y ticks with consistent decimal formatting."""
    numeric_values = pd.Series(values, dtype=float).replace(
        [float("inf"), float("-inf")],
        float("nan"),
    ).dropna()

    if numeric_values.empty:
        return fig

    data_min = float(numeric_values.min())
    data_max = float(numeric_values.max())

    if view == "growth":
        span = data_max - data_min
        if span == 0:
            span = max(abs(data_max) * 0.1, 1.0)
        axis_min = data_min - span * 0.05
        axis_max = data_max + span * 0.10
        tick_format = ".1f"
        tick_suffix = "%"
        tick_values = [
            axis_min + (axis_max - axis_min) * index / 4
            for index in range(5)
        ]
    else:
        axis_min = 0.0
        axis_max = float(max(int(data_max) + 1, 1))
        tick_format = ",g"
        tick_suffix = ""
        tick_values = [axis_max * index / 4 for index in range(5)]

    if axis_max <= axis_min:
        axis_max = axis_min + 1.0

    tick_text = [
        f"{value:{tick_format}}{tick_suffix}"
        for value in tick_values
    ]

    fig.update_yaxes(
        range=[axis_min, axis_max],
        tickmode="array",
        tickvals=tick_values,
        ticktext=tick_text,
    )
    return fig


def set_five_division_y_axis(fig, maximum, decimals=2):
    """Set a zero-based axis with five equal divisions (six tick marks)."""
    axis_max = max(float(maximum), 1.0)
    tick_values = [axis_max * index / 5 for index in range(6)]
    fig.update_yaxes(
        range=[0, axis_max],
        tickmode="array",
        tickvals=tick_values,
        ticktext=[f"{value:,.{decimals}f}" for value in tick_values],
    )
    return fig


def padded_whole_axis_ceiling(maximum):
    """Add 2 at the second-highest place, then round up to that place."""
    maximum = max(float(maximum), 0.0)
    if maximum == 0:
        return 1
    highest_place = 10 ** floor(log10(maximum))
    rounding_place = max(highest_place / 10, 1)
    padding = 2 * rounding_place
    return max(ceil((maximum + padding) / rounding_place) * rounding_place, 1)


# ============================================================
# 6. CHART 1: THE TREND
# ============================================================

def selected_years(start_year, end_year):
    return [year for year in YEARS if start_year <= year <= end_year]


def metric_axis_values(metric_code, values):
    """Return plotted values, unit, divisor, and display precision."""
    if metric_code == "PERCAPITA":
        return values / 1000, "thousand PhP per person", 1000, 2

    # Keep expenditure measures comparable and show larger numeric values.
    return values / 1_000, "PhP billion", 1_000, 1


def make_metric_levels(metric_code, start_year, end_year):
    years = selected_years(start_year, end_year)
    if not years:
        return make_empty_figure("No years selected.")

    metric_code = metric_code if metric_code in METRIC_SERIES else "THE"
    series = METRIC_SERIES[metric_code]
    values = series.reindex(years).fillna(0.0)
    plotted, unit, divisor, decimals = metric_axis_values(metric_code, values)
    fig = go.Figure(
        go.Scatter(
            x=years,
            y=plotted,
            mode="lines+markers+text",
            name=METRIC_LABELS[metric_code],
            line=dict(width=3, color="#a4cf95"),
            marker=dict(size=7, color="#a4cf95"),
            customdata=[f"{value / divisor:,.{decimals}f}" for value in values],
            text=[f"{value / divisor:,.{decimals}f}" for value in values],
            textposition="top center",
            textfont=dict(size=10, color="#f3f6fa"),
            cliponaxis=False,
            hovertemplate=(
                "Year: %{x}<br>"
                f"{METRIC_LABELS[metric_code]}: {unit} "
                "%{customdata}<extra></extra>"
            ),
        )
    )
    fig.update_layout(xaxis_title="", yaxis_title=unit, showlegend=False)
    fig.update_xaxes(
        tickmode="array", tickvals=years,
        ticktext=[str(year) for year in years], tickangle=-45,
    )
    fig = apply_chart_style(fig)
    axis_max = padded_whole_axis_ceiling(float(plotted.max()))
    fig = set_five_division_y_axis(fig, axis_max, decimals=0)
    return add_chart_heading(fig, METRIC_LABELS[metric_code], "Levels")


def make_metric_growth(metric_code, start_year, end_year):
    years = selected_years(start_year, end_year)
    if not years:
        return make_empty_figure("No years selected.")

    metric_code = metric_code if metric_code in METRIC_SERIES else "THE"
    series = METRIC_SERIES[metric_code].astype(float)
    # Calculate against the prior workbook year, including when the selected
    # range starts after the first available year.
    growth = series.pct_change(fill_method=None).mul(100).replace(
        [float("inf"), float("-inf")], float("nan")
    )
    plotted = growth.reindex(years)
    fig = go.Figure(
        go.Bar(
            x=years,
            y=plotted,
            name="Growth rate",
            marker_color=[
                "#F2B66D" if pd.notna(value) and value < 0 else "#4C9F38"
                for value in plotted
            ],
            text=["" if pd.isna(value) else f"{value:.1f}%" for value in plotted],
            textposition="outside",
            textfont=dict(size=10, color="#f3f6fa"),
            cliponaxis=False,
            hovertemplate=(
                "Year: %{x}<br>Growth rate: %{y:.1f}%<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        xaxis_title="", yaxis_title="Year-on-year growth",
        showlegend=False,
    )
    fig.update_xaxes(
        tickmode="array", tickvals=years,
        ticktext=[str(year) for year in years], tickangle=-45,
    )
    fig = apply_chart_style(fig)
    valid_growth = plotted.dropna()
    if not valid_growth.empty:
        low, high = float(valid_growth.min()), float(valid_growth.max())
        span = max(high - low, abs(high) * 0.1, abs(low) * 0.1, 1.0)
        axis_min = min(0.0, low - span * 0.08)
        axis_max = padded_whole_axis_ceiling(max(high, 0.0))
        if axis_max <= axis_min:
            axis_max = axis_min + 1.0
        ticks = [axis_min + (axis_max - axis_min) * i / 5 for i in range(6)]
        fig.update_yaxes(
            range=[axis_min, axis_max], tickmode="array", tickvals=ticks,
            ticktext=[f"{value:.0f}%" for value in ticks], zeroline=True,
            zerolinecolor="#94a3b8",
        )
    return add_chart_heading(fig, METRIC_LABELS[metric_code], "Growth Rates")


def make_che_stack_chart(classification_code, start_year, end_year, focused_year=None):
    years = selected_years(start_year, end_year)
    if not years:
        return make_empty_figure("No years selected.")
    if focused_year not in years:
        focused_year = None
    elif focused_year is not None:
        years = [focused_year]
    horizontal_focus = focused_year is not None
    if classification_code not in CHE_CODES:
        classification_code = "FS"

    category_tables = {
        year: get_raw_category_data(classification_code, year)
        for year in years
    }
    if all(frame.empty for frame in category_tables.values()):
        return make_empty_figure(
            f"No data available for {CLASSIFICATION_LABELS[classification_code]}."
        )

    totals = {}
    for frame in category_tables.values():
        for row in frame.itertuples(index=False):
            totals[row.Category] = totals.get(row.Category, 0.0) + float(row.Amount)
    ordered_categories = sorted(totals, key=totals.get, reverse=True)
    if classification_code == "REG":
        other_label = "Other Regions"
        named_regions = [
            category for category in ordered_categories
            if str(category).strip().lower() != "other regions"
        ]
        keep_categories = named_regions[:3]
        other_categories = set(ordered_categories) - set(keep_categories)
    else:
        other_label = "Other categories"
        keep_categories = ordered_categories[:TOP_CATEGORIES_TO_SHOW]
        focus_categories = [
            category for category in ordered_categories
            if is_focus_category(classification_code, category)
            or is_secondary_focus_category(classification_code, category)
        ]
        if focus_categories and focus_categories[0] not in keep_categories:
            keep_categories[-1] = focus_categories[0]
        other_categories = set(ordered_categories) - set(keep_categories)
        if "Other categories" in totals:
            other_categories.add("Other categories")
            keep_categories = [c for c in keep_categories if c != "Other categories"]

    matrix = {category: [] for category in keep_categories}
    if other_categories:
        matrix[other_label] = []
    for year in years:
        frame = category_tables[year]
        amounts = dict(zip(frame["Category"], frame["Amount"]))
        for category in keep_categories:
            matrix[category].append(float(amounts.get(category, 0.0)))
        if other_categories:
            matrix[other_label].append(
                sum(float(amounts.get(category, 0.0)) for category in other_categories)
            )

    fig = go.Figure()
    stack_palette = [
        "#1F5A32",
        "#2F7D32",
        "#4C9F38",
        "#6AAE59",
        "#86BF75",
        "#A4CF95",
        "#C3DEB8",
        "#E0EDD8",
    ]
    year_totals = [
        sum(amounts[index] for amounts in matrix.values())
        for index in range(len(years))
    ]
    share_annotations = []
    cumulative_shares = [0.0] * len(years)
    for category_index, (category, amounts) in enumerate(matrix.items()):
        primary_highlight = is_focus_category(classification_code, category)
        secondary_highlight = is_secondary_focus_category(classification_code, category)
        highlighted = primary_highlight or secondary_highlight
        shares = [
            amount / total * 100 if total > 0 else 0.0
            for amount, total in zip(amounts, year_totals)
        ]
        share_labels = [
            f"{share:.1f}%" if total > 0 else ""
            for share, total in zip(shares, year_totals)
        ] if highlighted else None
        trace_x = (
            shares
            if horizontal_focus
            else years
        )
        trace_y = (
            [str(year) for year in years]
            if horizontal_focus
            else shares
        )
        customdata = [
            (f"{share:.1f}%", f"PhP {amount / 1_000:,.2f} billion")
            for amount, share in zip(amounts, shares)
        ]
        fig.add_trace(
            go.Bar(
                x=trace_x,
                y=trace_y,
                orientation="h" if horizontal_focus else "v",
                name=category,
                marker_color=(
                    "#A6DDF7" if primary_highlight
                    else "#CBB8E8" if secondary_highlight
                    else stack_palette[category_index % len(stack_palette)]
                ),
                text=share_labels,
                textposition=(
                    "outside" if highlighted and horizontal_focus else "none"
                ),
                insidetextanchor="middle",
                textfont=dict(color="#14213D", size=11),
                customdata=customdata,
                hovertemplate=(
                    (f"Year: {focused_year}<br>" if horizontal_focus else "Year: %{x}<br>")
                    + f"{category}: %{{customdata[0]}}<br>"
                    + "Level: %{customdata[1]}<extra></extra>"
                ),
            )
        )
        if highlighted and not horizontal_focus:
            for index, (year, share, total) in enumerate(zip(years, shares, year_totals)):
                if total <= 0 or share <= 0:
                    continue
                share_annotations.append(
                    dict(
                        x=year + 0.36,
                        y=cumulative_shares[index] + share / 2,
                        xref="x",
                        yref="y",
                        text=f"{share:.1f}%",
                        showarrow=False,
                        xanchor="left",
                        yanchor="middle",
                        font=dict(color="#F3F7FF", size=10),
                        bgcolor="rgba(23, 32, 51, 0.72)",
                        borderpad=1,
                    )
                )
        cumulative_shares = [
            base + share for base, share in zip(cumulative_shares, shares)
        ]
    fig.update_layout(
        barmode="stack",
        xaxis_title="Share of CHE (%)" if horizontal_focus else "",
        yaxis_title="" if horizontal_focus else "Share of CHE (%)",
        bargap=0.30, showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.20, x=0),
        annotations=share_annotations,
    )
    fig = apply_chart_style(fig)
    share_ticks = [0, 20, 40, 60, 80, 100]
    if horizontal_focus:
        fig.update_xaxes(
            range=[0, 100], tickmode="array", tickvals=share_ticks,
            ticktext=[f"{value}%" for value in share_ticks],
        )
        fig.update_yaxes(
            tickmode="array", tickvals=[str(focused_year)],
            ticktext=[str(focused_year)],
        )
    else:
        fig.update_xaxes(
            tickmode="array", tickvals=years,
            ticktext=[str(year) for year in years], tickangle=-45,
            range=[years[0] - 0.5, years[-1] + 0.8],
        )
        fig.update_yaxes(
            range=[0, 100], tickmode="array", tickvals=share_ticks,
            ticktext=[f"{value}%" for value in share_ticks],
        )
    title = (
        "Current Health Expenditure by "
        f"{STACK_CLASSIFICATION_LABELS[classification_code]}"
    )
    subtitle = f"Share to total · {focused_year}" if horizontal_focus else "Share to total"
    fig = add_chart_heading(fig, title, subtitle, top_margin=64)
    # Reserve a separate band below the rotated year labels for the legend.
    fig.update_layout(
        margin=dict(l=35, r=25, t=64, b=96),
        legend=dict(orientation="h", yanchor="top", y=-0.30, xanchor="left", x=0),
    )
    return fig

def make_the_trend(start_year, end_year, view, selected_measure="THE"):
    years = [
        year for year in YEARS
        if start_year <= year <= end_year
    ]

    if not years:
        return make_empty_figure("No years selected.")

    series_by_measure = {
        "THE": THE_SERIES,
        "CHE": CHE_SERIES,
        "HK": HK_SERIES,
    }
    measure_labels = {
        "THE": "Total Health Expenditure",
        "CHE": "Current Health Expenditure",
        "HK": "Health Capital Formation",
    }
    series = series_by_measure.get(selected_measure, THE_SERIES)
    values = series.reindex(years).fillna(0.0)

    if view == "growth":
        plotted_values = values.pct_change() * 100
        plotted_values = plotted_values.replace(
            [float("inf"), float("-inf")],
            float("nan"),
        )

        y_title = "Year-on-year growth (%)"
        hover_template = (
            "Year: %{x}<br>"
            "Growth: %{y:.1f}%"
            "<extra></extra>"
        )
    else:
        plotted_values, amount_unit = scale_amounts_for_axis(values)
        amount_divisor = {
            "trillion": 1_000_000,
            "billion": 1_000,
            "million": 1,
        }[amount_unit]
        formatted_values = [
            f"PhP {value / amount_divisor:,.2f} {amount_unit}"
            for value in values
        ]
        y_title = f"Amount (PhP {amount_unit})"
        hover_template = (
            "Year: %{x}<br>"
            f"{selected_measure}: " + "%{customdata}<br>"
            "<extra></extra>"
        )

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=years,
            y=plotted_values,
            mode="lines+markers",
            name=selected_measure,
            line=dict(width=3, color="#172033"),
            marker=dict(size=7),
            hovertemplate=hover_template,
            customdata=(
                formatted_values
                if view != "growth" else None
            ),
        )
    )

    fig.update_layout(
        xaxis_title="",
        yaxis_title=y_title,
    )
    fig.update_xaxes(
        tickmode="array",
        tickvals=years,
        ticktext=[str(year) for year in years],
        tickangle=-45,
    )

    fig = apply_chart_style(fig)
    if view != "growth" and selected_measure == "THE":
        observed_max = float(plotted_values.max()) if len(plotted_values) else 0.0
        padding = 300 if amount_unit == "billion" else 0.3
        axis_max = max(
            round(observed_max + padding, 1),
            ceil(observed_max * 10) / 10,
            0.1,
        )
        fig = set_five_division_y_axis(fig, axis_max, decimals=1)
    else:
        fig = set_time_series_y_axis(fig, plotted_values, view)
    return add_chart_heading(
        fig,
        measure_labels.get(selected_measure, measure_labels["THE"]),
        "Growth Rate" if view == "growth" else "Levels",
    )


def make_special_trend_chart(
    title,
    series_list,
    start_year,
    end_year,
    view,
    unit,
    levels_axis_padding=None,
):
    years = [year for year in YEARS if start_year <= year <= end_year]
    if not years:
        return make_empty_figure("No years selected.")

    fig = go.Figure()
    plotted_series = []
    for label, source_series, divisor in series_list:
        values = source_series.reindex(years).fillna(0.0)
        if view == "growth":
            plotted = values.pct_change() * 100
            plotted = plotted.replace(
                [float("inf"), float("-inf")], float("nan")
            )
            hover = "Year: %{x}<br>Growth: %{y:.1f}%<extra>%{fullData.name}</extra>"
        else:
            plotted = values / divisor
            hover = (
                "Year: %{x}<br>"
                f"{label}: %{{y:,.2f}} {unit}"
                "<extra></extra>"
            )

        fig.add_trace(
            go.Scatter(
                x=years,
                y=plotted,
                mode="lines+markers",
                name=label,
                line=dict(width=3, color="#172033"),
                marker=dict(size=6),
                hovertemplate=hover,
            )
        )
        plotted_series.append(plotted)

    fig.update_layout(
        xaxis_title="",
        yaxis_title=(
            "Year-on-year growth (%)"
            if view == "growth"
            else f"Amount ({unit})"
        ),
        showlegend=False,
    )
    fig.update_xaxes(
        tickmode="array",
        tickvals=years,
        ticktext=[str(year) for year in years],
        tickangle=-45,
    )
    combined_values = pd.concat(plotted_series, ignore_index=True)
    fig = apply_chart_style(fig)
    if view == "growth":
        fig = set_time_series_y_axis(fig, combined_values, view)
    elif levels_axis_padding is not None:
        observed_max = float(combined_values.max()) if not combined_values.empty else 0.0
        fig = set_five_division_y_axis(
            fig,
            ceil(observed_max + levels_axis_padding),
        )
    else:
        fig = set_time_series_y_axis(fig, combined_values, view)
    return add_chart_heading(
        fig,
        title,
        "Growth Rate" if view == "growth" else "Levels",
        top_margin=68 if len(series_list) > 1 else 56,
    )


def make_phc_chart(start_year, end_year, view):
    years = [year for year in YEARS if start_year <= year <= end_year]
    if not years:
        return make_empty_figure("No years selected.")

    phc = PHC_SERIES.reindex(years).fillna(0.0)
    che = PHC_CHE_SERIES.reindex(years).fillna(0.0)
    if view == "growth":
        growth = (phc.pct_change() * 100).replace(
            [float("inf"), float("-inf")],
            float("nan"),
        )
        fig = go.Figure(
            go.Scatter(
                x=years,
                y=growth,
                mode="lines+markers",
                name="PHC",
                line=dict(width=3, color="#4C9F38"),
                marker=dict(size=7),
                hovertemplate=(
                    "Year: %{x}<br>PHC Growth: %{y:.1f}%"
                    "<extra></extra>"
                ),
            )
        )
        fig.update_layout(
            xaxis_title="",
            yaxis_title="Year-on-year growth (%)",
            showlegend=False,
        )
        fig.update_xaxes(
            tickmode="array",
            tickvals=years,
            ticktext=[str(year) for year in years],
            tickangle=-45,
        )
        fig = set_time_series_y_axis(apply_chart_style(fig), growth, "growth")
        return add_chart_heading(
            fig,
            "Primary Health Care Expenditure",
            "Growth Rate",
            top_margin=62,
        )

    che_scaled, amount_unit = scale_amounts_for_axis(che)
    divisor = {
        "trillion": 1_000_000,
        "billion": 1_000,
        "million": 1,
    }[amount_unit]
    phc_scaled = phc / divisor
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=years,
            y=che_scaled,
            width=0.72,
            name="CHE",
            marker_color="#A0C989",
            text=["CHE"] * len(years),
            textposition="outside",
                textfont=dict(size=9, color="#096a11"),
            customdata=[f"{value / divisor:,.2f}" for value in che],
            hovertemplate=(
                f"Year: %{{x}}<br>CHE: PhP %{{customdata}} {amount_unit}"
                "<extra></extra>"
            ),
        )
    )
    fig.add_trace(
        go.Bar(
            x=years,
            y=phc_scaled,
            width=0.72,
            name="PHC",
            marker_color="#4C9F38",
            text=["PHC"] * len(years),
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(size=9, color="white"),
            customdata=[f"{value / divisor:,.2f}" for value in phc],
            hovertemplate=(
                f"Year: %{{x}}<br>PHC: PhP %{{customdata}} {amount_unit}"
                "<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        barmode="overlay",
        showlegend=False,
        xaxis_title="",
        yaxis_title=f"Amount (PhP {amount_unit})",
        bargap=0.22,
    )
    fig.update_xaxes(
        tickmode="array",
        tickvals=years,
        ticktext=[str(year) for year in years],
        tickangle=-45,
    )
    fig = apply_chart_style(fig)
    observed_max = float(che_scaled.max()) if len(che_scaled) else 0.0
    padding = 300 if amount_unit == "billion" else 0.3
    axis_max = max(
        round(observed_max + padding, 1),
        ceil(observed_max * 10) / 10,
        0.1,
    )
    fig = set_five_division_y_axis(fig, axis_max, decimals=1)
    return add_chart_heading(
        fig,
        "Primary Health Care Expenditure",
        "Levels",
        top_margin=62,
    )


def make_percapita_chart(start_year, end_year, view):
    return make_special_trend_chart(
        "Current Health Expenditure Per Capita",
        [("Per capita expenditure", PERCAPITA_SERIES, 1000)],
        start_year,
        end_year,
        view,
        "thousand PhP per person",
        levels_axis_padding=3,
    )


# ============================================================
# 7. CHART 2: THE COMPOSITION
# ============================================================

def make_the_composition(year):
    che = float(CHE_SERIES.get(year, 0.0))
    hk = float(HK_SERIES.get(year, 0.0))

    if che + hk <= 0:
        return make_empty_figure(
            "No expenditure data available for this year."
        )

    fig = go.Figure(
        data=[
            go.Pie(
                labels=[
                    "Current Health Expenditure",
                    "Health Capital Formation",
                ],
                values=[che, hk],
                sort=False,
                rotation=(90 - 180 * che / (che + hk)) % 360,
                automargin=True,
                textinfo="label+percent",
                textposition="outside",
                textfont=dict(size=10, color="#096a11"),
                marker=dict(
                    colors=["#4C9F38", "#A0C989"],
                    line=dict(width=1, color="white"),
                ),
                hovertemplate=(
                    "<b>%{label}</b><br>"
                    "Amount: %{customdata}<br>"
                    "Share: %{percent}"
                    "<extra></extra>"
                ),
                customdata=[format_amount(che), format_amount(hk)],
            )
        ]
    )

    fig.update_layout(
        showlegend=False,
        uniformtext=dict(minsize=8, mode="show"),
        autosize=True,
        margin=dict(
            l=18,
            r=18,
            t=52,
            b=24,
        ),
        template="plotly_white",
        paper_bgcolor="white",
        font=dict(
            family="Arial, sans-serif",
            size=10,
            color="#096a11",
        ),
    )

    return add_chart_heading(
        fig,
        f"{year} Total Health Expenditure",
        "Share to total",
        top_margin=62,
    )


# ============================================================
# 8. CHART 3: SELECTED CHE / HK TREND
# ============================================================

def make_selected_trend(code, start_year, end_year, view):
    years = [
        year for year in YEARS
        if start_year <= year <= end_year
    ]

    if not years:
        return make_empty_figure("No years selected.")

    series = get_selected_series(code)
    values = series.reindex(years).fillna(0.0)

    if code == HK_CODE:
        series_title = "Health Capital Formation"
        series_name = "HK"
    else:
        series_title = "Current Health Expenditure"
        series_name = "CHE"

    if view == "growth":
        plotted_values = values.pct_change() * 100
        plotted_values = plotted_values.replace(
            [float("inf"), float("-inf")],
            float("nan"),
        )

        y_title = "Year-on-year growth (%)"
        hover_template = (
            "Year: %{x}<br>"
            "Growth: %{y:.1f}%"
            "<extra></extra>"
        )
    else:
        plotted_values, amount_unit = scale_amounts_for_axis(values)
        y_title = f"Amount (PhP {amount_unit})"
        hover_template = (
            "Year: %{x}<br>"
            f"{series_name}: "
            "%{customdata}<extra></extra>"
        )

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=years,
            y=plotted_values,
            mode="lines+markers",
            name=series_name,
            line=dict(width=3, color="#172033"),
            marker=dict(size=7),
            hovertemplate=hover_template,
            customdata=(
                [format_amount(value) for value in values]
                if view != "growth" else None
            ),
        )
    )

    fig.update_layout(
        xaxis_title="Year",
        yaxis_title=y_title,
    )
    fig.update_xaxes(
        tickmode="array",
        tickvals=years,
        ticktext=[str(year) for year in years],
        tickangle=-45,
    )

    fig = set_time_series_y_axis(apply_chart_style(fig), plotted_values, view)
    return add_chart_heading(
        fig,
        series_title,
        "Growth Rate" if view == "growth" else "Levels",
    )


# ============================================================
# 9. CLASSIFICATION TREEMAP
# ============================================================

def make_category_treemap(category_df, title, subtitle):
    scaled_values, amount_unit = scale_amounts_for_axis(
        category_df["Amount"]
    )
    category_colors = get_category_color_map(category_df)
    categories = category_df["Category"].astype(str).tolist()
    colors = [category_colors[category] for category in categories]
    tile_names = categories

    fig = go.Figure(
        data=[
            go.Treemap(
                labels=categories,
                parents=[""] * len(categories),
                values=scaled_values,
                text=tile_names,
                marker=dict(
                    colors=colors,
                    line=dict(width=2, color="#f5faf2"),
                ),
                sort=False,
                tiling=dict(packing="squarify", pad=3),
                texttemplate=(
                    "<b>%{text}</b>"
                    f"<br>PhP %{{value:,.2f}} {amount_unit}"
                    "<br>%{percentRoot:.1%}"
                ),
                textfont=dict(size=12, color="white"),
                hovertemplate=(
                    "<b>%{label}</b><br>"
                    "Amount: %{customdata}<br>"
                    "Share: %{percentRoot:.2%}"
                    "<extra></extra>"
                ),
                customdata=category_df["Amount"].map(format_amount),
                domain=dict(x=[0.02, 0.98], y=[0.02, 0.98]),
                pathbar=dict(visible=False),
            )
        ]
    )

    fig.update_layout(
        showlegend=False,
        uniformtext=dict(minsize=12, mode="show"),
        margin=dict(
            l=8,
            r=8,
            t=48,
            b=8,
        ),
        template="plotly_white",
        paper_bgcolor="white",
        font=dict(
            family="Arial, sans-serif",
            size=11,
            color="#096a11",
        ),
    )

    return add_chart_heading(fig, title, subtitle, top_margin=52)


def make_classification_treemap(code, year):
    if code == "THE":
        category_df = pd.DataFrame(
            [
                {"Category": "Current Health Expenditure", "Amount": float(CHE_SERIES.get(year, 0.0))},
                {"Category": "Health Capital Formation", "Amount": float(HK_SERIES.get(year, 0.0))},
            ]
        )
        return make_category_treemap(
            category_df,
            f"{year} Total Health Expenditure",
            "Levels and share to total",
        )

    if code == "CHE":
        code = "FS"
    category_df = get_category_data(code, year)

    if category_df.empty:
        return make_empty_figure(
            f"No data available for {code} in {year}."
        )

    if code == HK_CODE:
        title = "Health Capital Formation"
    else:
        title_label = CLASSIFICATION_LABELS.get(code, code)
        title = f"{year} Current Health Expenditure by {title_label}"
    return make_category_treemap(
        category_df,
        title,
        "Levels and share to total",
    )


# ============================================================
# 11. KEY FINDINGS
# ============================================================

def make_key_findings(code, year):
    che = float(CHE_SERIES.get(year, 0.0))
    hk = float(HK_SERIES.get(year, 0.0))
    the = float(THE_SERIES.get(year, 0.0))

    if year > YEARS[0]:
        previous_the = float(THE_SERIES.get(year - 1, 0.0))

        if previous_the != 0:
            growth = (
                (the - previous_the)
                / abs(previous_the)
                * 100
            )

            growth_text = (
                f"THE changed by {growth:,.2f}% "
                f"from {year - 1} to {year}."
            )
        else:
            growth_text = (
                f"THE for {year - 1} is zero or unavailable; "
                "growth cannot be calculated."
            )
    else:
        growth_text = (
            f"{year} is the first year in the displayed series."
        )

    che_share = che / the * 100 if the else 0.0
    hk_share = hk / the * 100 if the else 0.0

    category_df = get_category_data(code, year)

    if category_df.empty:
        largest_text = (
            f"No classification breakdown is available "
            f"for {code} in {year}."
        )
    else:
        largest = category_df.iloc[0]
        largest_text = (
            f"The largest displayed category for "
            f"{CLASSIFICATION_LABELS.get(code, code)} "
            f"is {largest['Category']}, with "
            f"{format_amount(float(largest['Amount']))}."
        )

    if code == "REG":
        grouping_text = (
            "The regional charts show the seven largest regions "
            "individually and combine the remaining regions "
            "into Other Regions."
        )
    else:
        grouping_text = (
            "The classification charts combine smaller categories "
            "into Other categories when necessary."
        )

    return [
        html.Div(
            [
                html.Div(
                    "THE composition",
                    className="finding-title",
                ),
                html.Div(
                    f"Current Health Expenditure represents "
                    f"{che_share:,.2f}% of THE, while "
                    f"Health Capital Formation represents "
                    f"{hk_share:,.2f}%.",
                    className="finding-text",
                ),
            ],
            className="finding-card",
        ),
        html.Div(
            [
                html.Div(
                    "Year-on-year movement",
                    className="finding-title",
                ),
                html.Div(
                    growth_text,
                    className="finding-text",
                ),
            ],
            className="finding-card",
        ),
        html.Div(
            [
                html.Div(
                    "Classification insight",
                    className="finding-title",
                ),
                html.Div(
                    largest_text,
                    className="finding-text",
                ),
            ],
            className="finding-card",
        ),
        html.Div(
            [
                html.Div(
                    "Chart grouping",
                    className="finding-title",
                ),
                html.Div(
                    grouping_text,
                    className="finding-text",
                ),
            ],
            className="finding-card",
        ),
    ]


# ============================================================
# 12. DASH APP AND CSS
# ============================================================

app = Dash(
    __name__,
    assets_folder=str(Path(__file__).resolve().parent / "assets"),
    title="Philippine National Health Account Dashboard",
    suppress_callback_exceptions=True,
)

app.index_string = """
<!DOCTYPE html>
<html>
<head>
    {%metas%}
    <title>{%title%}</title>
    {%favicon%}
    {%css%}
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
        * {
            box-sizing: border-box;
        }

        :root {
            --theme-blue: #14213D;
            --theme-green: #4C9F38;
            --theme-tertiary: #096a11;
            --theme-font: #172033;
            --theme-border: #d4deea;
            --theme-surface: #ffffff;
            --theme-wash: #1b2b47;
        }

        body {
            margin: 0;
            background: var(--theme-wash);
            color: var(--theme-font);
            font-family: Arial, sans-serif;
            height: 100vh;
            overflow: hidden;
        }

        .app-shell {
            display: flex;
            height: 100vh;
            height: 100dvh;
            width: 100%;
            overflow: hidden;
        }

        /* COMPACT SIDEBAR */
        .sidebar {
            flex: 0 0 20%;
            width: 20%;
            min-width: 205px;
            max-width: none;
            background: var(--theme-blue);
            color: white;
            padding: 12px 9px;
            display: flex;
            flex-direction: column;
            gap: 7px;
            height: 100vh;
            height: 100dvh;
            overflow: hidden;
        }

        .brand {
            display: flex;
            align-items: center;
            gap: 9px;
            min-width: 0;
            padding: 3px 6px 9px;
            border-bottom: 1px solid #344156;
        }

        .brand-icon {
            width: 48px;
            height: 48px;
            flex: 0 0 48px;
            object-fit: contain;
        }

        .brand-title {
            font-size: 16px;
            font-weight: 700;
            line-height: 1.3;
        }

        .sidebar-section-title {
            font-size: 12px;
            font-weight: 600;
            color: #d9e2f0;
            letter-spacing: 0;
            margin: 3px 5px;
            line-height: 1.35;
        }

        .sidebar-static-section {
            margin-top: 12px;
            padding-top: 10px;
            border-top: 1px solid #344156;
        }

        .sidebar-static-item {
            color: #d9e2f0;
            font-size: 12px;
            font-weight: 600;
            padding: 7px 8px;
            line-height: 1.25;
        }

        .nav-container {
            display: flex;
            flex-direction: column;
            gap: 2px;
            overflow: hidden;
            min-height: 0;
        }

        .che-nav-container {
            gap: 2px;
        }

        .nav-button {
            width: 100%;
            border: 0;
            border-radius: 5px;
            background: transparent;
            color: #d9e2f0;
            text-align: left;
            padding: 7px 8px;
            font-size: 12px;
            font-weight: 600;
            line-height: 1.25;
            cursor: pointer;
        }

        .nav-main {
            font-size: 12px;
            font-weight: 600;
        }

        .nav-button:hover {
            background: transparent;
            color: #d9e2f0;
        }

        .nav-button.active {
            background: var(--theme-green);
            color: white;
            font-weight: 700;
        }

        .nav-subitem {
            display: flex;
            flex: 0 0 42px !important;
            align-items: center;
            height: 42px !important;
            min-height: 42px !important;
            max-height: 42px;
            padding: 3px 8px 3px 18px;
            font-size: 11px;
            line-height: 1.2;
        }

        .che-nav-container .nav-subitem {
            flex-basis: 32px !important;
            height: 32px !important;
            min-height: 32px !important;
            max-height: 32px;
            padding-top: 2px;
            padding-bottom: 2px;
        }

        .nav-subitem > span {
            display: block;
            width: 100%;
        }

        .classification-definition {
            flex: 1 1 auto;
            min-height: 120px;
            margin-top: 5px;
            padding: 12px;
            overflow: hidden;
            border: 1px solid #49617f;
            border-radius: 8px;
            background: #203452;
            color: #e8eef8;
        }

        .classification-definition-title {
            margin-bottom: 9px;
            color: #a4cf95;
            font-size: 18px;
            font-weight: 700;
            letter-spacing: 0.04em;
            line-height: 1.3;
        }

        .classification-definition-text {
            font-size: calc(12px + 3pt);
            line-height: 1.5;
            text-align: justify;
        }

        .sidebar-controls {
            border-top: 1px solid #344156;
            padding-top: 9px;
            display: flex;
            flex-direction: column;
            gap: 8px;
            margin-top: auto;
        }

        .control-label {
            color: #d9e2f0;
            font-size: 11px;
            margin-bottom: 4px;
            display: block;
        }

        .control-label-row {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 7px;
        }

        .year-control-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 7px;
        }

        .control-block {
            min-width: 0;
        }

        /* Force dropdown text to black on white */
        .sidebar .Select-control,
        .sidebar .Select-menu-outer,
        .sidebar .Select-menu,
        .sidebar .Select-option,
        .sidebar .Select-value,
        .sidebar .Select-placeholder,
        .sidebar .Select-value-label,
        .sidebar .Select-input,
        .sidebar .Select-input input {
            color: #111827 !important;
            background-color: #ffffff !important;
        }

        .sidebar .Select-control {
            border: 1px solid var(--theme-border) !important;
            border-radius: 5px !important;
            min-height: 31px !important;
        }

        .sidebar .Select-value-label {
            color: #111827 !important;
        }

        .sidebar .Select-placeholder {
            color: #64748b !important;
        }

        .sidebar .Select-menu-outer {
            z-index: 9999;
        }

        /* Newer Dash dropdown class names */
        .sidebar .dash-dropdown,
        .sidebar .Select,
        .sidebar .VirtualizedSelectOption {
            color: #111827;
        }

        .sidebar .Select-option:hover {
            background-color: #e7f3e2 !important;
        }

        .sidebar .Select-option.is-focused {
            background-color: #e7f3e2 !important;
            color: #111827 !important;
        }

        .sidebar .Select-option.is-selected {
            background-color: #e7f3e2 !important;
            color: #111827 !important;
        }

        /* MAIN CONTENT */
        .main-content {
            flex: 1;
            min-width: 0;
            min-height: 0;
            height: 100vh;
            height: 100dvh;
            padding: 17px;
            background: linear-gradient(115deg, #1b2b47 0%, #203657 52%, #243d60 100%);
            display: grid;
            grid-template-rows: auto minmax(0, 1fr) auto;
            overflow: hidden;
        }

        .page-header {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 12px;
            flex-wrap: wrap;
            margin-bottom: 14px;
            flex: 0 0 auto;
        }

        .page-title {
            margin: 0;
            font-size: 23px;
            font-weight: 700;
            color: var(--theme-font);
        }

        .page-subtitle {
            color: #4d7d47;
            font-size: 12px;
            margin-top: 6px;
        }

        .kpi-carousel {
            display: flex;
            align-items: stretch;
            gap: 8px;
            min-width: 0;
            margin-bottom: 14px;
            flex: 0 0 auto;
        }

        .kpi-track {
            position: relative;
            flex: 1 1 auto;
            min-width: 0;
            height: 106px;
            overflow: hidden;
        }

        .kpi-slot-0 { left: calc(0% + 5px); }
        .kpi-slot-1 { left: calc(25% + 5px); }
        .kpi-slot-2 { left: calc(50% + 5px); }
        .kpi-slot-3 { left: calc(75% + 5px); }
        .kpi-slot-4 { left: calc(100% + 5px); }
        .kpi-slot-5 { left: calc(125% + 5px); }
        .kpi-slot-6 { left: calc(150% + 5px); }
        .kpi-slot-7 { left: calc(175% + 5px); }
        .kpi-slot-8 { left: calc(200% + 5px); }
        .kpi-slot-9 { left: calc(225% + 5px); }

        .kpi-card {
            position: absolute;
            top: 0;
            box-sizing: border-box;
            width: calc(25% - 10px);
            height: 100%;
            transition: left 0.45s ease, transform 0.2s ease, box-shadow 0.2s ease;
        }

        .kpi-card:hover {
            transform: scale(1.05);
            z-index: 2;
            box-shadow: 0 7px 20px rgba(76, 159, 56, 0.2);
        }

        .kpi-no-transition {
            transition: none !important;
        }

        .kpi-arrow {
            flex: 0 0 34px;
            width: 34px;
            height: 34px;
            align-self: center;
            border: 1px solid var(--theme-border);
            border-radius: 50%;
            background: white;
            color: var(--theme-green);
            font-size: 22px;
            line-height: 1;
            cursor: pointer;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.08);
        }

        .kpi-arrow:hover {
            background: #edf7e9;
            border-color: var(--theme-green);
        }

        .kpi-card {
            background: linear-gradient(145deg, #2a4568 0%, #355476 100%);
            border: 1px solid #567494;
            border-top: 3px solid var(--theme-green);
            border-left: 3px solid #a4cf95;
            border-radius: 12px;
            padding: 8px 10px;
            min-width: 0;
            box-shadow: 0 4px 12px rgba(76, 159, 56, 0.08);
            display: flex;
            flex-direction: column;
            justify-content: center;
        }

        #kpi-the-card,
        #kpi-che-card,
        #kpi-hk-card,
        #kpi-phc-card,
        #kpi-percapita-card,
        #kpi-gov-card,
        #kpi-oop-card,
        #kpi-hospitals-card,
        #kpi-curative-card,
        #kpi-ncr-card {
            padding-left: 82px;
            padding-right: 10px;
        }

        #kpi-the-card::after,
        #kpi-che-card::after,
        #kpi-hk-card::after,
        #kpi-phc-card::after,
        #kpi-percapita-card::after,
        #kpi-gov-card::after,
        #kpi-oop-card::after,
        #kpi-hospitals-card::after,
        #kpi-curative-card::after,
        #kpi-ncr-card::after {
            content: "";
            position: absolute;
            top: 50%;
            left: 11px;
            width: 60px;
            height: 60px;
            transform: translateY(-50%);
            background-position: center;
            background-repeat: no-repeat;
            background-size: contain;
            pointer-events: none;
        }

        #kpi-the-card::after { background-image: url("/assets/KPI1-THE.png"); }
        #kpi-che-card::after { background-image: url("/assets/KPI2-CHE.png"); }
        #kpi-hk-card::after { background-image: url("/assets/KPI3-HK.png"); }
        #kpi-phc-card::after { background-image: url("/assets/KPI4-PHC.png"); }
        #kpi-percapita-card::after { background-image: url("/assets/KPI5-PERCAPITA.png"); }
        #kpi-gov-card::after { background-image: url("/assets/KPI6%20-%20Government%20FSRI.png"); }
        #kpi-oop-card::after { background-image: url("/assets/KPI7%20-%20Household%20HF.png"); }
        #kpi-hospitals-card::after { background-image: url("/assets/KPI8%20-%20Hospitals%20HP.png"); }
        #kpi-curative-card::after { background-image: url("/assets/KPI9%20-%20Curative%20HC.png"); }
        #kpi-ncr-card::after { background-image: url("/assets/KPI10%20-%20NCR%20Region.png"); }

        .kpi-heading {
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 6px;
            min-width: 0;
            margin-bottom: 5px;
        }

        .kpi-label {
            color: #a4cf95;
            font-size: 11px;
            font-weight: 600;
            margin: 0;
            line-height: 1.1;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .kpi-value {
            font-size: 20px;
            font-weight: 800;
            color: #ffffff;
            overflow-wrap: anywhere;
            line-height: 1.1;
        }

        .kpi-secondary-value {
            margin-top: 2px;
            color: #d9e2f0;
            font-size: 17px;
            font-weight: 700;
            line-height: 1.1;
        }

        .kpi-description {
            color: #d9e2f0;
            font-size: 10px;
            line-height: 1.2;
            margin-top: 4px;
            text-align: left;
            white-space: nowrap;
            flex: 0 0 auto;
        }

        .dashboard-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            grid-template-rows: minmax(0, 1fr) minmax(0, 1.25fr);
            grid-template-areas:
                "levels growth"
                "che-stack che-stack";
            gap: 12px;
            align-items: stretch;
            min-height: 0;
        }

        .chart-levels { grid-area: levels; }
        .chart-growth { grid-area: growth; }
        .chart-che-stack {
            grid-area: che-stack;
            display: flex;
            flex-direction: column;
            min-height: 0;
            padding-bottom: 12px;
        }

        .stack-chart-plot {
            flex: 1 1 auto;
            min-height: 0;
        }

        .chart-card {
            background: #2a4568;
            border: 1px solid #567494;
            border-top: 3px solid var(--theme-green);
            border-left: 3px solid #a4cf95;
            border-radius: 12px;
            min-width: 0;
            overflow: hidden;
            padding: 2px;
            box-shadow: 0 4px 14px rgba(76, 159, 56, 0.08);
            transition: box-shadow 0.18s ease, transform 0.18s ease;
        }

        .stack-analysis-text {
            display: flex;
            align-items: center;
            min-width: 0;
            min-height: 53px;
            flex: 0 0 auto;
            margin: 0 12px 0;
            padding: 10px 12px;
            color: #f3f7ff;
            background: transparent;
            border: 1px solid #567494;
            border-left: 4px solid #a6ddf7;
            border-radius: 7px;
            font-size: calc(12px + 4pt);
            line-height: 1.4;
        }

        .chart-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(76, 159, 56, 0.14);
        }

        .chart-top {
            min-height: 0;
        }

        .chart-treemap {
            min-height: 0;
        }

        .findings-panel {
            display: flex;
            flex-direction: column;
            min-height: 0;
            padding: 10px;
        }

        .findings-panel .section-heading {
            margin: 0 0 6px;
        }

        .findings-panel .findings-grid {
            flex: 1 1 auto;
            min-height: 0;
            margin-top: 0;
        }

        .chart-card .js-plotly-plot,
        .chart-card .plot-container {
            width: 100% !important;
            height: 100% !important;
        }

        .section-heading {
            font-size: 15px;
            font-weight: 700;
            margin: 19px 0 9px;
            color: var(--theme-font);
            flex: 0 0 auto;
        }

        .findings-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 10px;
            margin-top: 10px;
            flex: 0 0 auto;
        }

        .finding-card {
            background: white;
            border: 1px solid var(--theme-border);
            border-left: 3px solid var(--theme-green);
            border-radius: 10px;
            padding: 11px 13px;
            min-width: 0;
        }

        .finding-title {
            font-size: 12px;
            font-weight: 700;
            margin-bottom: 6px;
            color: var(--theme-font);
        }

        .finding-text {
            font-size: 11px;
            line-height: 1.5;
            color: #52627a;
        }

        .footer {
            color: #d9e2f0;
            font-size: 10px;
            margin-top: 18px;
            padding-bottom: 6px;
            flex: 0 0 auto;
        }

        .footer a {
            color: #a4cf95;
            font-weight: 600;
            text-decoration: underline;
            text-underline-offset: 2px;
        }

        .footer a:hover {
            color: #e0edd8;
        }

        @media (max-width: 1350px) {
            .main-content {
                padding: 13px;
            }

            .kpi-value {
                font-size: 17px;
            }

            .chart-top {
                min-height: 0;
            }
        }

        @media (max-width: 1050px) {
            .app-shell {
                flex-direction: column;
            }

            .sidebar {
                width: 100%;
                max-width: none;
                min-width: 0;
                height: auto;
                max-height: 34dvh;
                flex: 0 0 auto;
            }

            .nav-container {
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
            }

            .sidebar-controls {
                margin-top: 0;
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }

        }

        @media (max-width: 700px) {
            .main-content {
                height: calc(100dvh - 34dvh);
                padding: 10px;
            }

            .nav-container {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }

            .sidebar-controls {
                grid-template-columns: 1fr;
            }

            .dashboard-grid {
                grid-template-columns: minmax(0, 1fr);
                grid-template-rows: repeat(3, minmax(110px, 1fr));
                grid-template-areas: "levels" "growth" "che-stack";
            }

            .kpi-card {
                padding: 6px 8px;
            }

            .kpi-value {
                font-size: 15px;
            }

            .findings-grid {
                grid-template-columns: 1fr;
                display: none;
            }

            .section-heading,
            .footer {
                display: none;
            }

            .page-title {
                font-size: 19px;
            }
        }

    </style>
</head>
<body>
    {%app_entry%}
    <footer>
        {%config%}
        {%scripts%}
        {%renderer%}
    </footer>
</body>
</html>
"""

# Resolve icon URLs through Dash so they keep working when the app is hosted
# under a prefixed path, such as Posit Connect Cloud.
KPI_ICON_FILES = [
    "KPI1-THE.png",
    "KPI2-CHE.png",
    "KPI3-HK.png",
    "KPI4-PHC.png",
    "KPI5-PERCAPITA.png",
    "KPI6 - Government FSRI.png",
    "KPI7 - Household HF.png",
    "KPI8 - Hospitals HP.png",
    "KPI9 - Curative HC.png",
    "KPI10 - NCR Region.png",
]
for icon_file in KPI_ICON_FILES:
    encoded_file = icon_file.replace(" ", "%20")
    app.index_string = app.index_string.replace(
        f'url("/assets/{encoded_file}")',
        f'url("{app.get_asset_url(icon_file)}")',
    )


# ============================================================
# 13. SIDEBAR NAVIGATION
# ============================================================

def make_nav_button(code):
    label = f"by {CLASSIFICATION_LABELS.get(code, code)}"

    return html.Button(
        html.Span(label),
        id={
            "type": "che-classification-button",
            "code": code,
        },
        n_clicks=0,
        className="nav-button nav-subitem",
        title=code,
    )


def make_metric_button(code):
    return html.Button(
        METRIC_LABELS[code],
        id={"type": "metric-button", "code": code},
        n_clicks=0,
        className="nav-button nav-main",
        title=code,
    )


# ============================================================
# 14. MAIN LAYOUT
# ============================================================

app.layout = html.Div(
    className="app-shell",
    children=[
        # ---------------- SIDEBAR ----------------
        html.Aside(
            className="sidebar",
            children=[
                html.Div(
                    [
                        html.Img(
                            src=app.get_asset_url("PNHA Icon.png"),
                            alt="Philippine National Health Accounts logo",
                            className="brand-icon",
                        ),
                        html.Div(
                            [
                                "Philippine National Health",
                                html.Br(),
                                "Account Dashboard",
                            ],
                            className="brand-title",
                        ),
                    ],
                    className="brand",
                ),

                html.Div(
                    [
                        make_metric_button(code)
                        for code in ("THE", "CHE", "PERCAPITA", "HK", "PHC")
                    ],
                    className="nav-container",
                ),
                html.Div(
                    "CURRENT HEALTH EXPENDITURE",
                    className="sidebar-static-section",
                ),
                html.Div(
                    [make_nav_button(code) for code in CHE_CODES],
                    className="nav-container che-nav-container",
                ),
                html.Div(
                    [
                        html.Div(
                            CHE_CLASSIFICATION_DEFINITION_TITLES["FS"],
                            id="che-classification-definition-title",
                            className="classification-definition-title",
                        ),
                        html.Div(
                            CHE_CLASSIFICATION_DEFINITIONS["FS"],
                            id="che-classification-definition",
                            className="classification-definition-text",
                        ),
                    ],
                    className="classification-definition",
                ),

                # Sidebar controls in the requested order.
                html.Div(
                    className="sidebar-controls",
                    children=[
                        # Shared year range for all three time-series charts.
                        html.Div(
                            [
                                html.Div(
                                    [
                                        html.Label(
                                            "Start year",
                                            className="control-label",
                                        ),
                                        html.Label(
                                            "End year",
                                            className="control-label",
                                        ),
                                    ],
                                    className="control-label-row",
                                ),

                                html.Div(
                                    [
                                        dcc.Dropdown(
                                            id="start-year",
                                            options=[
                                                {
                                                    "label": str(year),
                                                    "value": year,
                                                }
                                                for year in YEARS
                                            ],
                                            value=2014,
                                            clearable=False,
                                            searchable=False,
                                        ),

                                        dcc.Dropdown(
                                            id="end-year",
                                            options=[
                                                {
                                                    "label": str(year),
                                                    "value": year,
                                                }
                                                for year in YEARS
                                            ],
                                            value=YEARS[-1],
                                            clearable=False,
                                            searchable=False,
                                        ),
                                    ],
                                    className="year-control-grid",
                                ),
                            ],
                            className="control-block",
                        ),

                    ],
                ),
            ],
        ),

        # ---------------- MAIN DASHBOARD ----------------
        html.Main(
            className="main-content",
            children=[
                dcc.Store(
                    id="selected-metric",
                    data="THE",
                    storage_type="memory",
                ),
                dcc.Store(
                    id="selected-che-classification",
                    data="FS",
                    storage_type="memory",
                ),
                dcc.Store(
                    id="che-stack-focused-year",
                    data=None,
                    storage_type="memory",
                ),
                dcc.Store(
                    id="kpi-carousel-position",
                    data=0,
                    storage_type="memory",
                ),

                # KPI CARDS
                html.Div(
                    className="kpi-carousel",
                    children=[
                        html.Button(
                            "‹",
                            id="kpi-carousel-prev",
                            n_clicks=0,
                            className="kpi-arrow",
                            title="Rotate KPI cards left",
                        ),
                        html.Div(
                            className="kpi-track",
                            children=[
                                html.Div(
                                    id="kpi-the-card",
                                    className="kpi-card kpi-slot-0",
                                    children=[
                                        html.Div("TOTAL HEALTH EXPENDITURE", className="kpi-label"),
                                        html.Div(id="kpi-the", className="kpi-value"),
                                        html.Div(id="kpi-the-year", className="kpi-description"),
                                    ],
                                ),
                                html.Div(
                                    id="kpi-che-card",
                                    className="kpi-card kpi-slot-1",
                                    children=[
                                        html.Div("CURRENT HEALTH EXPENDITURE", className="kpi-label"),
                                        html.Div(id="kpi-che", className="kpi-value"),
                                        html.Div(id="kpi-che-year", className="kpi-description"),
                                    ],
                                ),
                                html.Div(
                                    id="kpi-hk-card",
                                    className="kpi-card kpi-slot-2",
                                    children=[
                                        html.Div("HEALTH CAPITAL FORMATION", className="kpi-label"),
                                        html.Div(id="kpi-hk", className="kpi-value"),
                                        html.Div(id="kpi-hk-year", className="kpi-description"),
                                    ],
                                ),
                                html.Div(
                                    id="kpi-phc-card",
                                    className="kpi-card kpi-slot-3",
                                    children=[
                                        html.Div("PRIMARY HEALTH CARE EXPENDITURE", className="kpi-label"),
                                        html.Div(id="kpi-phc", className="kpi-value"),
                                        html.Div(id="kpi-phc-year", className="kpi-description"),
                                    ],
                                ),
                                html.Div(
                                    id="kpi-percapita-card",
                                    className="kpi-card kpi-slot-4",
                                    children=[
                                        html.Div("CHE PER CAPITA", className="kpi-label"),
                                        html.Div(id="kpi-percapita", className="kpi-value"),
                                        html.Div(id="kpi-percapita-year", className="kpi-description"),
                                    ],
                                ),
                                html.Div(
                                    id="kpi-gov-card",
                                    className="kpi-card kpi-slot-5",
                                    children=[
                                        html.Div(
                                            html.Div("GOVERNMENT SHARE TO CHE", className="kpi-label"),
                                            className="kpi-heading",
                                        ),
                                        html.Div(id="kpi-gov-level", className="kpi-value"),
                                        html.Div(id="kpi-gov-share", className="kpi-secondary-value"),
                                        html.Div(id="kpi-gov-year", className="kpi-description"),
                                    ],
                                ),
                                html.Div(
                                    id="kpi-oop-card",
                                    className="kpi-card kpi-slot-6",
                                    children=[
                                        html.Div(
                                            html.Div("HOUSEHOLD OOP SHARE TO CHE", className="kpi-label"),
                                            className="kpi-heading",
                                        ),
                                        html.Div(id="kpi-oop-level", className="kpi-value"),
                                        html.Div(id="kpi-oop-share", className="kpi-secondary-value"),
                                        html.Div(id="kpi-oop-year", className="kpi-description"),
                                    ],
                                ),
                                html.Div(
                                    id="kpi-hospitals-card",
                                    className="kpi-card kpi-slot-7",
                                    children=[
                                        html.Div(
                                            html.Div("HOSPITALS SHARE TO CHE", className="kpi-label"),
                                            className="kpi-heading",
                                        ),
                                        html.Div(id="kpi-hospitals-level", className="kpi-value"),
                                        html.Div(id="kpi-hospitals-share", className="kpi-secondary-value"),
                                        html.Div(id="kpi-hospitals-year", className="kpi-description"),
                                    ],
                                ),
                                html.Div(
                                    id="kpi-curative-card",
                                    className="kpi-card kpi-slot-8",
                                    children=[
                                        html.Div(
                                            html.Div("CURATIVE CARE SHARE TO CHE", className="kpi-label"),
                                            className="kpi-heading",
                                        ),
                                        html.Div(id="kpi-curative-level", className="kpi-value"),
                                        html.Div(id="kpi-curative-share", className="kpi-secondary-value"),
                                        html.Div(id="kpi-curative-year", className="kpi-description"),
                                    ],
                                ),
                                html.Div(
                                    id="kpi-ncr-card",
                                    className="kpi-card kpi-slot-9",
                                    children=[
                                        html.Div(
                                            html.Div("NCR SHARE TO CHE", className="kpi-label"),
                                            className="kpi-heading",
                                        ),
                                        html.Div(id="kpi-ncr-level", className="kpi-value"),
                                        html.Div(id="kpi-ncr-share", className="kpi-secondary-value"),
                                        html.Div(id="kpi-ncr-year", className="kpi-description"),
                                    ],
                                ),
                            ],
                        ),
                        html.Button(
                            "›",
                            id="kpi-carousel-next",
                            n_clicks=0,
                            className="kpi-arrow",
                            title="Rotate KPI cards right",
                        ),
                    ],
                ),

                html.Div(
                    className="dashboard-grid",
                    children=[
                        html.Div(
                            dcc.Graph(
                                id="metric-levels-chart",
                                config={"displayModeBar": False},
                                responsive=True,
                                style={"height": "100%"},
                            ),
                            className="chart-card chart-levels",
                        ),
                        html.Div(
                            dcc.Graph(
                                id="metric-growth-chart",
                                config={"displayModeBar": False},
                                responsive=True,
                                style={"height": "100%"},
                            ),
                            className="chart-card chart-growth",
                        ),
                        html.Div(
                            children=[
                                html.Div(
                                    dcc.Graph(
                                        id="che-stack-chart",
                                        config={"displayModeBar": False},
                                        responsive=True,
                                        style={"height": "100%"},
                                    ),
                                    className="stack-chart-plot",
                                ),
                                html.Div(
                                    id="che-analysis-text",
                                    className="stack-analysis-text",
                                ),
                            ],
                            className="chart-card chart-che-stack",
                        ),
                    ],
                ),

                html.Div(
                    [
                        "Sources: Philippine Statistics Authority, 2025 Philippine National Health Accounts ",
                        html.A(
                            "Statistical Tables",
                            href="https://psa.gov.ph/statistics/pnha/tables",
                            target="_blank",
                            rel="noopener noreferrer",
                        ),
                        " · ",
                        html.A(
                            "Technical Notes",
                            href="https://psa.gov.ph/statistics/pnha/technical-notes",
                            target="_blank",
                            rel="noopener noreferrer",
                        ),
                        ".",
                        html.Br(),
                        "Coverage: 2014–2025. Values are in Philippine pesos.",
                        html.Br(),
                        "Prepared by: E-conomist · Last Update: October 19, 2026",
                    ],
                    className="footer",
                ),
            ],
        ),
    ],
)


# ============================================================
# 15. CALLBACKS: SIDEBAR SELECTIONS
# ============================================================

@app.callback(
    Output("selected-metric", "data"),
    Input(
        {
            "type": "metric-button",
            "code": ALL,
        },
        "n_clicks",
    ),
    prevent_initial_call=True,
)
def update_selected_metric(n_clicks):
    triggered = ctx.triggered_id
    return triggered["code"] if isinstance(triggered, dict) else no_update


@app.callback(
    Output("selected-che-classification", "data"),
    Input(
        {
            "type": "che-classification-button",
            "code": ALL,
        },
        "n_clicks",
    ),
    prevent_initial_call=True,
)
def update_selected_che_classification(n_clicks):
    triggered = ctx.triggered_id
    return triggered["code"] if isinstance(triggered, dict) else no_update


@app.callback(
    Output("che-classification-definition-title", "children"),
    Output("che-classification-definition", "children"),
    Input("selected-che-classification", "data"),
)
def update_che_classification_definition(classification_code):
    if classification_code not in CHE_CLASSIFICATION_DEFINITIONS:
        classification_code = "FS"
    return (
        CHE_CLASSIFICATION_DEFINITION_TITLES[classification_code],
        CHE_CLASSIFICATION_DEFINITIONS[classification_code],
    )


@app.callback(
    Output(
        {
            "type": "metric-button",
            "code": ALL,
        },
        "className",
    ),
    Input("selected-metric", "data"),
)
def update_metric_nav_styles(selected_metric):
    return [
        "nav-button nav-main active" if code == selected_metric
        else "nav-button nav-main"
        for code in ("THE", "CHE", "PERCAPITA", "HK", "PHC")
    ]


@app.callback(
    Output(
        {"type": "che-classification-button", "code": ALL},
        "className",
    ),
    Input("selected-che-classification", "data"),
)
def update_che_nav_styles(selected_classification):
    return [
        "nav-button nav-subitem active" if code == selected_classification
        else "nav-button nav-subitem"
        for code in CHE_CODES
    ]


# ============================================================
# 17. CALLBACKS: VALIDATE YEAR RANGE
# ============================================================

@app.callback(
    Output("start-year", "value"),
    Output("end-year", "value"),
    Input("start-year", "value"),
    Input("end-year", "value"),
    prevent_initial_call=True,
)
def keep_year_range_valid(start_year, end_year):
    if start_year is None or end_year is None or start_year <= end_year:
        return no_update, no_update

    if ctx.triggered_id == "start-year":
        return no_update, start_year

    return end_year, no_update


@app.callback(
    Output("che-stack-focused-year", "data"),
    Input("che-stack-chart", "clickData"),
    State("che-stack-focused-year", "data"),
    State("start-year", "value"),
    State("end-year", "value"),
    prevent_initial_call=True,
)
def toggle_che_stack_year(click_data, focused_year, start_year, end_year):
    if not click_data or not click_data.get("points"):
        return no_update

    if (
        focused_year is not None
        and start_year is not None
        and end_year is not None
        and start_year <= focused_year <= end_year
    ):
        return None

    clicked_value = click_data["points"][0].get("x")
    try:
        clicked_year = int(float(clicked_value))
    except (TypeError, ValueError):
        return no_update

    if start_year is not None and clicked_year < start_year:
        return no_update
    if end_year is not None and clicked_year > end_year:
        return no_update
    return clicked_year if clicked_year in YEARS else no_update


# ============================================================
# 18. CALLBACK: UPDATE DASHBOARD
# ============================================================

@app.callback(
    Output("kpi-the", "children"),
    Output("kpi-che", "children"),
    Output("kpi-hk", "children"),
    Output("kpi-phc", "children"),
    Output("kpi-percapita", "children"),
    Output("kpi-the-year", "children"),
    Output("kpi-che-year", "children"),
    Output("kpi-hk-year", "children"),
    Output("kpi-phc-year", "children"),
    Output("kpi-percapita-year", "children"),
    Output("kpi-gov-share", "children"),
    Output("kpi-gov-level", "children"),
    Output("kpi-gov-year", "children"),
    Output("kpi-oop-share", "children"),
    Output("kpi-oop-level", "children"),
    Output("kpi-oop-year", "children"),
    Output("kpi-hospitals-share", "children"),
    Output("kpi-hospitals-level", "children"),
    Output("kpi-hospitals-year", "children"),
    Output("kpi-curative-share", "children"),
    Output("kpi-curative-level", "children"),
    Output("kpi-curative-year", "children"),
    Output("kpi-ncr-share", "children"),
    Output("kpi-ncr-level", "children"),
    Output("kpi-ncr-year", "children"),
    Output("metric-levels-chart", "figure"),
    Output("metric-growth-chart", "figure"),
    Output("che-stack-chart", "figure"),
    Output("che-analysis-text", "children"),
    Input("selected-metric", "data"),
    Input("selected-che-classification", "data"),
    Input("start-year", "value"),
    Input("end-year", "value"),
    Input("che-stack-focused-year", "data"),
)
def update_dashboard(
    selected_metric,
    selected_classification,
    start_year,
    end_year,
    focused_stack_year,
):
    if selected_metric not in METRIC_SERIES:
        selected_metric = "THE"
    if selected_classification not in CHE_CODES:
        selected_classification = "FS"

    if start_year is None:
        start_year = YEARS[0]

    if end_year is None:
        end_year = YEARS[-1]

    if start_year > end_year:
        start_year, end_year = end_year, start_year

    selected_year = end_year

    if focused_stack_year not in selected_years(start_year, end_year):
        focused_stack_year = None

    # KPI VALUES
    che = float(CHE_SERIES.get(selected_year, 0.0))
    hk = float(HK_SERIES.get(selected_year, 0.0))
    the = float(THE_SERIES.get(selected_year, 0.0))

    focus_kpi_categories = [
        ("FSRI", ["Government"]),
        ("HF", CHE_ANALYSIS_ALIASES["HF"]),
        ("HP", ["Hospitals", "Hospital"]),
        ("HC", ["Curative care", "Curative"]),
        ("REG", ["NCR", "National Capital Region"]),
    ]
    focus_kpi_values = []
    for sheet_name, aliases in focus_kpi_categories:
        amount = get_named_category_amount(sheet_name, selected_year, aliases)
        che_total = float(CHE_SERIES.get(selected_year, 0.0))
        share = amount / che_total * 100 if che_total else 0.0
        focus_kpi_values.extend([
            f"Share of CHE: {share:.1f}%",
            format_kpi_amount(amount),
            str(selected_year),
        ])

    levels_fig = make_metric_levels(selected_metric, start_year, end_year)
    growth_fig = make_metric_growth(selected_metric, start_year, end_year)
    che_stack_fig = make_che_stack_chart(
        selected_classification,
        start_year,
        end_year,
        focused_stack_year,
    )

    return (
        format_kpi_amount(the),
        format_kpi_amount(che),
        format_kpi_amount(hk),
        format_kpi_amount(float(PHC_SERIES.get(selected_year, 0.0))),
        f"PhP {float(PERCAPITA_SERIES.get(selected_year, 0.0)):,.2f}",
        str(selected_year),
        str(selected_year),
        str(selected_year),
        str(selected_year),
        str(selected_year),
        *focus_kpi_values,
        levels_fig,
        growth_fig,
        che_stack_fig,
        make_che_analysis_text(selected_classification, start_year, end_year),
    )


@app.callback(
    Output("kpi-the-card", "className"),
    Output("kpi-che-card", "className"),
    Output("kpi-hk-card", "className"),
    Output("kpi-phc-card", "className"),
    Output("kpi-percapita-card", "className"),
    Output("kpi-gov-card", "className"),
    Output("kpi-oop-card", "className"),
    Output("kpi-hospitals-card", "className"),
    Output("kpi-curative-card", "className"),
    Output("kpi-ncr-card", "className"),
    Output("kpi-carousel-position", "data"),
    Input("kpi-carousel-prev", "n_clicks"),
    Input("kpi-carousel-next", "n_clicks"),
    State("kpi-carousel-position", "data"),
    prevent_initial_call=True,
)
def rotate_kpi_cards(previous_clicks, next_clicks, position):
    card_count = 10
    old_position = int(position or 0) % card_count
    position = old_position
    if ctx.triggered_id == "kpi-carousel-prev":
        position = (position - 1) % card_count
        wrapped_card = position
    elif ctx.triggered_id == "kpi-carousel-next":
        position = (position + 1) % card_count
        wrapped_card = old_position
    else:
        wrapped_card = None

    classes = [
        "kpi-card kpi-slot-" + str((card_index - position) % card_count)
        + (" kpi-no-transition" if card_index == wrapped_card else "")
        for card_index in range(card_count)
    ]
    return (*classes, position)


# ============================================================
# 19. RUN APP
# ============================================================

if __name__ == "__main__":
    app.run(
        debug=True,
        host="0.0.0.0",
        port=8050,
    )
