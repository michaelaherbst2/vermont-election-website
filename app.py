# To run:
# uv run streamlit run app.py

from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import base64
import html
import json
import asyncio
import threading
import time

from monitor import check_results
import pandas as pd
import streamlit as st


# =========================================================
# SETTINGS
# =========================================================
NEW_YORK_TZ = ZoneInfo(
    "America/New_York"
)

APP_DIR = Path(__file__).resolve().parent

RESULTS_DIR = APP_DIR / "vermont_results"

FEDERAL_DIR = RESULTS_DIR / "federal"

STATEWIDE_DIR = RESULTS_DIR / "statewide"
STATEWIDE_DEM_DIR = STATEWIDE_DIR / "democrat"
STATEWIDE_REP_DIR = STATEWIDE_DIR / "republican"

FEDERAL_DEM_FILE = (
    FEDERAL_DIR / "democratic_results.csv"
)

FEDERAL_REP_FILE = (
    FEDERAL_DIR / "republican_results.csv"
)

STATUS_FILE = RESULTS_DIR / "status.json"

LOGO = APP_DIR / "ap_elections.png"

DEFAULT_TOTAL_UNITS = 247


# =========================================================
# ELECTION DISPLAY
# =========================================================

ELECTION_NAME = "2026 Vermont Primary Election"
ELECTION_DATE_DISPLAY = "August 11, 2026"

EASTERN = ZoneInfo("America/New_York")


# =========================================================
# STATEWIDE OFFICES
# =========================================================

STATEWIDE_OFFICES = [
    "GOVERNOR",
    "LIEUTENANT GOVERNOR",
    "STATE TREASURER",
    "SECRETARY OF STATE",
    "AUDITOR OF ACCOUNTS",
    "ATTORNEY GENERAL",
]


STATEWIDE_FILES = {
    "GOVERNOR": {
        "dem": "dem_governor.xlsx",
        "rep": "rep_governor.xlsx",
    },

    "LIEUTENANT GOVERNOR": {
        "dem": "dem_lieutenant_governor.xlsx",
        "rep": "rep_lieutenant_governor.xlsx",
    },

    "STATE TREASURER": {
        "dem": "dem_state_treasurer.xlsx",
        "rep": "rep_state_treasurer.xlsx",
    },

    "SECRETARY OF STATE": {
        "dem": "dem_secretary_of_state.xlsx",
        "rep": "rep_secretary_of_state.xlsx",
    },

    "AUDITOR OF ACCOUNTS": {
        "dem": "dem_auditor_of_accounts.xlsx",
        "rep": "rep_auditor_of_accounts.xlsx",
    },

    "ATTORNEY GENERAL": {
        "dem": "dem_attorney_general.xlsx",
        "rep": "rep_attorney_general.xlsx",
    },
}


# =========================================================
# DIRECTORIES
# =========================================================

RESULTS_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

FEDERAL_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

STATEWIDE_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

STATEWIDE_DEM_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

STATEWIDE_REP_DIR.mkdir(
    parents=True,
    exist_ok=True,
)


# =========================================================
# PAGE
# =========================================================

st.set_page_config(
    page_title="Vermont Election Results",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =========================================================
# SESSION STATE
# =========================================================

defaults = {
    "federal_dem_full": False,
    "federal_rep_full": False,
    "statewide_dem_full": False,
    "statewide_rep_full": False,
}

for key, value in defaults.items():

    if key not in st.session_state:

        st.session_state[key] = value


# =========================================================
# NAVIGATION
# =========================================================

section = st.query_params.get(
    "section",
    "federal",
).lower()

if section not in {
    "federal",
    "statewide",
}:

    section = "federal"


# =========================================================
# CSS
# =========================================================

st.html(
    """
<style>

html,
body,
[data-testid="stAppViewContainer"],
.stApp {
    background:
        radial-gradient(
            circle at 55% 15%,
            #111516 0%,
            #080a0a 28%,
            #030404 58%,
            #000000 100%
        ) !important;

    color: #ffffff !important;
}

[data-testid="stHeader"] {
    background: transparent !important;
}

[data-testid="stToolbar"] {
    display: none !important;
}

[data-testid="stSidebar"] {
    display: none !important;
}

.block-container {
    max-width: 100% !important;
    padding: 0 0 40px 0 !important;
}


/* =====================================================
   HEADER
   ===================================================== */

.top-header {
    height: 98px;
    padding: 0 64px;

    display: flex;
    align-items: center;
    justify-content: space-between;

    border-bottom: 1px solid #252525;

    background: #020303;
}

.top-header-left {
    display: flex;
    align-items: center;
    gap: 22px;
}

.ap-logo {
    width: 56px;
    height: auto;
    display: block;
}

.ap-logo-fallback {
    width: 54px;
    height: 54px;

    background: white;
    color: black;

    font-size: 30px;
    font-weight: 900;

    display: flex;
    align-items: center;
    justify-content: center;

    position: relative;
}

.ap-logo-fallback::after {
    content: "";

    position: absolute;

    left: 0;
    right: 0;
    bottom: -7px;

    height: 7px;

    background: #ff2933;
}

.ap-header-title {
    font-size: 27px;
    font-weight: 700;
    line-height: 1.05;
}

.ap-header-subtitle {
    font-size: 14px;
    color: #cccccc;
    margin-top: 5px;
}

.ap-header-status {
    font-size: 13px;
    color: #eeeeee;
    white-space: nowrap;
}

.ap-unofficial {
    color: #ff4048;
    font-weight: 800;
}

.ap-result-status {
    color: #2fda61;
    font-weight: 800;
}

.ap-pipe {
    color: #666666;
    margin: 0 16px;
}


/* =====================================================
   LEFT NAV
   ===================================================== */

.nav-box {
    min-height: calc(100vh - 98px);

    border-right: 1px solid #252525;

    padding:
        38px
        20px
        35px
        22px;

    background:
        linear-gradient(
            180deg,
            #080909,
            #030404
        );
}

.nav-title {
    font-size: 19px;
    font-weight: 700;
    margin-bottom: 19px;
}

.nav-link {
    text-decoration: none !important;
    color: #eeeeee !important;
}

.nav-row {
    height: 43px;

    padding: 0 15px;
    margin-bottom: 4px;

    display: flex;
    align-items: center;

    border-radius: 6px;

    font-size: 15px;

    gap: 11px;

    cursor: pointer;
}

.nav-row:hover {
    background: #141414;
}

.nav-active {
    background: #181818;
    font-weight: 700;
}

.dot-red {
    color: #ff4248;
    font-size: 19px;
}

.dot-white {
    color: #f2f2f2;
    font-size: 18px;
}

.nav-rule {
    margin: 31px 3px 26px 3px;

    border-top: 1px solid #333333;
}

.nav-stat {
    color: #bfbfbf;
    font-size: 14px;
    margin-bottom: 19px;
}

.nav-stat strong {
    color: white;
}

.nav-bottom-label {
    color: #aaaaaa;
    font-size: 12px;
    margin-bottom: 5px;
}

.nav-bottom-value {
    color: #ffffff;
    font-size: 13px;
    margin-bottom: 24px;
}

.green-dot {
    color: #2fda61;
}


/* =====================================================
   MAIN
   ===================================================== */

.main-title-row {
    display: flex;
    align-items: center;
    gap: 8px;
}

.main-title {
    font-size: 32px;
    font-weight: 700;
    line-height: 1.05;
}

.link-icon {
    color: #5d7e93;
    font-size: 18px;
}

.main-subtitle {
    font-size: 15px;

    margin-top: 11px;
    margin-bottom: 18px;
}


/* =====================================================
   METRIC CARDS
   ===================================================== */

.metric-card {
    min-height: 116px;

    border: 1px solid #373737;
    border-radius: 7px;

    padding: 18px 21px;

    background:
        linear-gradient(
            135deg,
            rgba(23,23,23,.92),
            rgba(8,9,9,.92)
        );
}

.metric-inner {
    display: flex;
    align-items: center;
    gap: 18px;
}

.status-circle {
    width: 42px;
    height: 42px;

    border-radius: 50%;

    display: flex;
    align-items: center;
    justify-content: center;

    font-size: 25px;

    flex: 0 0 auto;
}

.status-circle-dem {
    border: 2px solid #4187ff;
    color: #4187ff;
}

.status-circle-rep {
    border: 2px solid #ff4148;
    color: #ff4148;
}

.status-circle-neutral {
    border: 2px solid #777777;
    color: #dddddd;
    font-size: 20px;
}

.metric-label {
    font-size: 14px;
    margin-bottom: 6px;
}

.metric-value {
    font-size: 27px;
    font-weight: 700;
    line-height: 1;
    margin-bottom: 8px;
}

.metric-dem {
    color: #478cff;
}

.metric-rep {
    color: #ff4148;
}

.metric-sub {
    color: #eeeeee;
    font-size: 14px;
}


/* =====================================================
   STATUS
   ===================================================== */

.status-line {
    font-size: 13px;
    color: #bcbcbc;

    margin:
        17px
        0
        25px
        0;
}


/* =====================================================
   TABS
   ===================================================== */

.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    border-bottom: 1px solid #252525;
}

.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border: none !important;

    color: #eeeeee !important;

    font-size: 14px !important;
    font-weight: 600 !important;

    padding:
        8px
        12px
        11px
        12px !important;
}

.stTabs [aria-selected="true"] {
    color: #438cff !important;
}

.stTabs [data-baseweb="tab-highlight"] {
    background-color: #438cff !important;
}


/* =====================================================
   OFFICE SELECTOR
   ===================================================== */

.office-label {
    color: #bbbbbb;
    font-size: 12px;
    margin-top: 4px;
    margin-bottom: 5px;
}

[data-testid="stSelectbox"] {
    max-width: 430px;
}

/* Closed selectbox */
[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    background: #080909 !important;
    border-color: #3a3a3a !important;
    color: #ffffff !important;
}

/* Selected office text */
[data-testid="stSelectbox"] div[data-baseweb="select"] span {
    color: #ffffff !important;
}

/* Dropdown arrow */
/* Dropdown arrow — black */
[data-testid="stSelectbox"] svg {
    fill: #000000 !important;
    color: #000000 !important;
}

/* Open dropdown */
div[data-baseweb="popover"] {
    background: #080909 !important;
}

div[data-baseweb="popover"] > div {
    background: #080909 !important;
}

/* Dropdown list */
ul[role="listbox"] {
    background: #080909 !important;
}

/* Every office option */
li[role="option"] {
    background: #080909 !important;
    color: #ffffff !important;
}

/* Text inside options */
li[role="option"] * {
    color: #ffffff !important;
}

/* Hovered option */
li[role="option"]:hover {
    background: #181818 !important;
    color: #ffffff !important;
}

/* Currently selected option */
li[role="option"][aria-selected="true"] {
    background: #202020 !important;
    color: #ffffff !important;
}

/* Keep selected option text white */
li[role="option"][aria-selected="true"] * {
    color: #ffffff !important;
}

/* =====================================================
   RESULT CARD
   ===================================================== */

.result-card {
    border: 1px solid #303030;

    border-radius: 7px;

    background:
        linear-gradient(
            135deg,
            rgba(11,13,13,.96),
            rgba(5,6,6,.96)
        );

    padding:
        19px
        20px
        15px
        20px;

    margin-bottom: 14px;
}

.result-heading {
    font-size: 25px;
    font-weight: 600;
}

.result-heading-dem {
    color: #4288ff;
}

.result-heading-rep {
    color: #ff4148;
}

.result-ru {
    font-size: 14px;
    margin-top: 6px;
}


/* =====================================================
   SEARCH
   ===================================================== */

[data-testid="stTextInput"] {
    margin: 0 0 10px 0 !important;
}

[data-testid="stTextInput"] > div > div {
    background: #080909 !important;

    border: 1px solid #393939 !important;

    border-radius: 5px !important;

    min-height: 40px !important;
}

[data-testid="stTextInput"] input {
    background: transparent !important;

    color: #ffffff !important;

    font-size: 13px !important;
}

[data-testid="stTextInput"] input::placeholder {
    color: #b6b6b6 !important;
}


/* =====================================================
   TABLE
   ===================================================== */

.results-table-wrap {
    width: 100%;

    overflow-x: auto;
    overflow-y: auto;

    max-height: 650px;

    margin-top: 8px;

    border-bottom: 1px solid #343434;
}

.results-table {
    width: 100%;

    border-collapse: separate;
    border-spacing: 0;

    font-size: 13px;

    border: 1px solid #343434;

    table-layout: auto;
}

.results-table thead th {
    position: sticky;
    top: 0;
    z-index: 20;
}

.results-table th {
    background:
        linear-gradient(
            90deg,
            #173568,
            #213f70
        );

    color: #f5f5f5;

    text-align: left;

    padding: 10px 12px;

    border-right:
        1px solid rgba(
            255,
            255,
            255,
            .08
        );

    border-bottom: 1px solid #343434;

    font-weight: 600;

    white-space: nowrap;
}

.results-table.rep-table th {
    background:
        linear-gradient(
            90deg,
            #681b20,
            #752329
        );
}

.results-table td {
    padding: 9px 12px;

    border-top: 1px solid #323232;

    border-right: 1px solid #282828;

    background: rgba(5,6,6,.98);

    color: white;

    white-space: nowrap;
}

.results-table td:not(:first-child) {
    text-align: right;
}

.results-table td:nth-child(1),
.results-table td:nth-child(2),
.results-table td:nth-child(3),
.results-table td:nth-child(4) {
    text-align: left;
}


/* =====================================================
   COLUMN WIDTHS
   ===================================================== */

/* Last Updated */
.results-table th:nth-child(1),
.results-table td:nth-child(1) {
    width: 220px;
    min-width: 220px;
    max-width: 220px;
    text-align: left;
    white-space: nowrap;
}

/* Town */
.results-table th:nth-child(2),
.results-table td:nth-child(2) {
    width: 165px;
    min-width: 145px;
    max-width: 180px;
    text-align: left;
    white-space: nowrap;
}

/* Rep District */
.results-table th:nth-child(3),
.results-table td:nth-child(3) {
    width: 185px;
    min-width: 150px;
    max-width: 210px;
    text-align: left;
    white-space: normal;
    overflow-wrap: break-word;
}

/* Sen District */
.results-table th:nth-child(4),
.results-table td:nth-child(4) {
    width: 185px;
    min-width: 150px;
    max-width: 220px;
    text-align: left;
    white-space: normal;
    overflow-wrap: break-word;
}


/* =====================================================
   BUTTON
   ===================================================== */

div[data-testid="stButton"] {
    display: flex;

    justify-content: center;

    margin-top: 12px;
}

div[data-testid="stButton"] button {
    background: transparent !important;

    border: 1px solid #3f7fe8 !important;

    color: #4b8fff !important;

    border-radius: 5px !important;

    padding: 7px 16px !important;

    font-size: 13px !important;

    font-weight: 600 !important;
}



/* =====================================================
   STATEWIDE OFFICE SELECT — WHITE OUTLINE + WHITE ARROW
   ===================================================== */

/* Closed select box */
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    background: #050606 !important;
    border: 1px solid #ffffff !important;
    color: #ffffff !important;
    box-shadow: none !important;
}

/* Keep white outline while focused/open */
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:focus-within {
    border: 1px solid #ffffff !important;
    box-shadow: none !important;
}

/* Selected office text */
div[data-testid="stSelectbox"] div[data-baseweb="select"] span,
div[data-testid="stSelectbox"] div[data-baseweb="select"] input {
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
}

/* Dropdown arrow */
div[data-testid="stSelectbox"] div[data-baseweb="select"] svg {
    fill: #ffffff !important;
    color: #ffffff !important;
}

/* Open dropdown/popover outer box */
div[data-baseweb="popover"] > div {
    background: #050606 !important;
    border: 1px solid #ffffff !important;
    box-shadow: none !important;
}

/* Open dropdown menu */
div[data-baseweb="menu"] {
    background: #050606 !important;
    border: 1px solid #ffffff !important;
    border-radius: 8px !important;
}

/* Dropdown option text */
div[data-baseweb="menu"] li,
div[data-baseweb="menu"] li span {
    color: #ffffff !important;
}

/* Selected / hovered option */
div[data-baseweb="menu"] li[aria-selected="true"],
div[data-baseweb="menu"] li:hover {
    background: #181818 !important;
    color: #ffffff !important;
}

</style>
"""
)


# =========================================================
# LOADERS
# =========================================================

def load_excel(path):

    if not path.exists():

        return None

    try:

        return pd.read_excel(
            path
        )

    except Exception as error:

        st.error(
            f"Could not read "
            f"{path.name}: "
            f"{error}"
        )

        return None


def load_csv(path):

    if not path.exists():

        return None

    try:

        return pd.read_csv(
            path,
            dtype=str,
        )

    except Exception as error:

        st.error(
            f"Could not read "
            f"{path.name}: "
            f"{error}"
        )

        return None


def load_status():

    if not STATUS_FILE.exists():

        return {}

    try:

        return json.loads(
            STATUS_FILE.read_text()
        )

    except Exception:

        return {}


def image_to_base64(path):

    if not path.exists():

        return None

    return base64.b64encode(
        path.read_bytes()
    ).decode(
        "utf-8"
    )


# =========================================================
# REPORTING
# =========================================================

def reporting_towns(df):

    if (
        df is None
        or "Town" not in df.columns
    ):

        return set()


    if "Total" in df.columns:

        total_column = "Total"


    elif "Total Votes" in df.columns:

        total_column = "Total Votes"


    else:

        return set()


    totals = (
        pd.to_numeric(
            df[
                total_column
            ],
            errors="coerce",
        )
        .fillna(0)
    )


    towns = (
        df.loc[
            totals > 0,
            "Town",
        ]
        .astype(str)
        .str.strip()
        .str.upper()
    )


    return set(
        towns
    )


def reporting_count_single(df):

    return len(
        reporting_towns(
            df
        )
    )


def get_reporting_status(
    status,
    dem=None,
    rep=None,
    prefix="federal",
):

    total_units = int(
        status.get(
            f"{prefix}_total_units",
            DEFAULT_TOTAL_UNITS,
        )
    )


    dem_fallback = (
        reporting_count_single(
            dem
        )
        if dem is not None
        else 0
    )


    rep_fallback = (
        reporting_count_single(
            rep
        )
        if rep is not None
        else 0
    )


    dem_units = int(
        status.get(
            f"{prefix}_dem_reporting",
            dem_fallback,
        )
    )


    rep_units = int(
        status.get(
            f"{prefix}_rep_reporting",
            rep_fallback,
        )
    )


    overall_units = int(
        status.get(
            f"{prefix}_overall_reporting",
            max(
                dem_units,
                rep_units,
            ),
        )
    )


    return (
        dem_units,
        rep_units,
        overall_units,
        total_units,
    )


# =========================================================
# SEARCH
# =========================================================

def filter_results(
    df,
    search,
):

    if df is None:

        return None


    term = (
        search
        .strip()
        .casefold()
    )


    if not term:

        return df


    if "Town" in df.columns:

        town_mask = (
            df["Town"]
            .astype(str)
            .str.casefold()
            .str.contains(
                term,
                regex=False,
            )
        )


        if town_mask.any():

            return df.loc[
                town_mask
            ]


    matching_columns = [
        column
        for column in df.columns
        if term in str(
            column
        ).casefold()
    ]


    if matching_columns:

        keep = []


        for column in [
            "Town",
            "Rep District",
            "RepDistrict",
            "Sen District",
            "SenDistrict",
        ]:

            if column in df.columns:

                keep.append(
                    column
                )


        keep += matching_columns


        for column in [
            "Total Write Ins",
            "Others",
            "Total",
            "Total Votes",
        ]:

            if column in df.columns:

                keep.append(
                    column
                )


        return df[
            list(
                dict.fromkeys(
                    keep
                )
            )
        ]


    mask = (
        df.astype(str)
        .apply(
            lambda row:
            row
            .str.casefold()
            .str.contains(
                term,
                regex=False,
            )
            .any(),
            axis=1,
        )
    )


    return df.loc[
        mask
    ]


# =========================================================
# LATEST TIMESTAMP
# =========================================================

def latest_timestamp(values):

    cleaned = []

    for value in values:

        if pd.isna(value):
            continue

        text = str(
            value
        ).strip()

        if not text:
            continue

        cleaned.append(
            text
        )

    if not cleaned:
        return ""

    parsed = pd.to_datetime(
        pd.Series(
            cleaned
        ),
        errors="coerce",
    )

    if parsed.notna().any():

        latest_index = (
            parsed
            .idxmax()
        )

        return cleaned[
            latest_index
        ]

    return cleaned[-1]


# =========================================================
# UNIQUE TEXT
# =========================================================

def combine_unique(values):

    cleaned = []


    for value in values:

        if pd.isna(value):

            continue


        text = str(
            value
        ).strip()


        if not text:

            continue


        if text.casefold() == "nan":

            continue


        if text not in cleaned:

            cleaned.append(
                text
            )


    return ", ".join(
        cleaned
    )


# =========================================================
# CLEAN RESULTS
# =========================================================

def clean_results(
    df,
    keep_districts=True,
):

    if df is None:

        return None


    display = df.copy()


    display.columns = [
        str(column).strip()
        for column in display.columns
    ]


    display = display.rename(
        columns={
            "RepDistrict":
                "Rep District",

            "SenDistrict":
                "Sen District",
        }
    )


    if "Town" not in display.columns:

        return display


    display["Town"] = (
        display["Town"]
        .astype(str)
        .str.strip()
        .str.upper()
    )


    # =====================================================
    # DISTRICTS
    # =====================================================

    if not keep_districts:

        display = display.drop(
            columns=[
                column
                for column in [
                    "Rep District",
                    "Sen District",
                ]
                if column in display.columns
            ]
        )


    # =====================================================
    # DROP UNWANTED
    # =====================================================

    display = display.drop(
        columns=[
            column
            for column in [
                "Overvotes",
                "Blank Votes",
            ]
            if column in display.columns
        ]
    )


    # =====================================================
    # TOTAL WRITE INS
    #
    # IMPORTANT:
    # Do NOT include the already-created
    # "Total Write Ins" field in write_in_columns.
    # =====================================================

    existing_total_write_ins = None


    if "Total Write Ins" in display.columns:

        existing_total_write_ins = (
            pd.to_numeric(
                display[
                    "Total Write Ins"
                ],
                errors="coerce",
            )
            .fillna(0)
        )


    write_in_columns = [
        column
        for column in display.columns
        if (
            column != "Total Write Ins"

            and

            (
                "write-in"
                in str(
                    column
                ).casefold()

                or

                "write in"
                in str(
                    column
                ).casefold()
            )
        )
    ]


    if write_in_columns:

        calculated_write_ins = pd.Series(
            0,
            index=display.index,
            dtype="float64",
        )


        for column in write_in_columns:

            calculated_write_ins += (
                pd.to_numeric(
                    display[
                        column
                    ],
                    errors="coerce",
                )
                .fillna(0)
            )


        if existing_total_write_ins is not None:

            calculated_write_ins += (
                existing_total_write_ins
            )


        display[
            "Total Write Ins"
        ] = (
            calculated_write_ins
            .fillna(0)
            .astype(int)
        )


        display = display.drop(
            columns=
                write_in_columns
        )


    elif existing_total_write_ins is not None:

        display[
            "Total Write Ins"
        ] = (
            existing_total_write_ins
            .fillna(0)
            .astype(int)
        )


    else:

        display[
            "Total Write Ins"
        ] = 0


    # =====================================================
    # OTHERS
    # =====================================================

    if "Others" in display.columns:

        display[
            "Others"
        ] = (
            pd.to_numeric(
                display[
                    "Others"
                ],
                errors="coerce",
            )
            .fillna(0)
            .astype(int)
        )


    # =====================================================
    # TOTAL
    # =====================================================

    if (
        "Total"
        not in display.columns

        and

        "Total Votes"
        in display.columns
    ):

        display = display.rename(
            columns={
                "Total Votes":
                    "Total"
            }
        )


    elif (
        "Total"
        in display.columns

        and

        "Total Votes"
        in display.columns
    ):

        display = display.drop(
            columns=[
                "Total Votes"
            ]
        )


    # =====================================================
    # REMOVE ZERO-ONLY CANDIDATES
    # =====================================================

    protected = {
        "Last Updated",
        "Town",
        "Rep District",
        "Sen District",
        "Total Write Ins",
        "Others",
        "Total",
    }


    drop_columns = []


    for column in display.columns:

        if column in protected:

            continue


        numeric = pd.to_numeric(
            display[
                column
            ],
            errors="coerce",
        )


        if (
            numeric.notna().any()

            and

            numeric
            .fillna(0)
            .sum()
            == 0
        ):

            drop_columns.append(
                column
            )


    if drop_columns:

        display = display.drop(
            columns=
                drop_columns
        )


    # =====================================================
    # NUMERIC COLUMNS
    # =====================================================

    text_columns = {
        "Last Updated",
        "Town",
        "Rep District",
        "Sen District",
    }


    vote_columns = []


    for column in display.columns:

        if column in text_columns:

            continue


        numeric = pd.to_numeric(
            display[
                column
            ],
            errors="coerce",
        )


        if numeric.notna().any():

            display[
                column
            ] = (
                numeric
                .fillna(0)
            )


            vote_columns.append(
                column
            )


    # =====================================================
    # COMBINE DUPLICATE TOWNS
    # =====================================================

    aggregation = {}


    for column in display.columns:

        if column == "Town":

            continue


        if column == "Last Updated":

            aggregation[
                column
            ] = latest_timestamp

        elif column in {
            "Rep District",
            "Sen District",
        }:

            aggregation[
                column
            ] = combine_unique


        elif column in vote_columns:

            aggregation[
                column
            ] = "sum"


        else:

            aggregation[
                column
            ] = combine_unique


    display = (
        display
        .groupby(
            "Town",
            as_index=False,
        )
        .agg(
            aggregation
        )
    )


    # =====================================================
    # INTEGER FORMAT
    # =====================================================

    for column in vote_columns:

        if column in display.columns:

            display[
                column
            ] = (
                pd.to_numeric(
                    display[
                        column
                    ],
                    errors="coerce",
                )
                .fillna(0)
                .astype(int)
            )


    # =====================================================
    # COLUMN ORDER
    # =====================================================

    first_columns = [
        column
        for column in [
            "Last Updated",
            "Town",
            "Rep District",
            "Sen District",
        ]
        if column in display.columns
    ]


    near_end = [
        column
        for column in [
            "Total Write Ins",
            "Others",
        ]
        if column in display.columns
    ]


    candidate_columns = [
        column
        for column in display.columns
        if (
            column not in first_columns

            and

            column not in near_end

            and

            column != "Total"
        )
    ]


    ordered = (
        first_columns
        +
        candidate_columns
        +
        near_end
    )


    if "Total" in display.columns:

        ordered.append(
            "Total"
        )


    display = display[
        ordered
    ]


    # =====================================================
    # MOST RECENT ACTUAL CHANGES FIRST
    # =====================================================

    if "Last Updated" in display.columns:

        display[
            "_last_updated_sort"
        ] = pd.to_datetime(
            display[
                "Last Updated"
            ],
            errors="coerce",
        )

        display = (
            display
            .sort_values(
                [
                    "_last_updated_sort",
                    "Town",
                ],
                ascending=[
                    False,
                    True,
                ],
                na_position="last",
            )
            .drop(
                columns=[
                    "_last_updated_sort"
                ]
            )
            .reset_index(
                drop=True
            )
        )

    else:

        display = (
            display
            .sort_values(
                "Town"
            )
            .reset_index(
                drop=True
            )
        )


    return display


# =========================================================
# EASTERN TIME DISPLAY
# =========================================================

def format_eastern_timestamp(value):

    """
    Display timestamps in Eastern Time.

    Important:
    - Existing strings like "08/12/2026 04:59:14 PM"
      are already Eastern and must NOT be converted again.
    - ISO/UTC strings like "2026-08-12T20:59:14Z"
      are converted to America/New_York.
    """

    if value is None:
        return "—"

    text = str(
        value
    ).strip()

    if (
        not text
        or text == "—"
        or text.casefold() == "nan"
    ):
        return "—"


    # -------------------------------------------------
    # ALREADY EASTERN DISPLAY FORMAT
    # -------------------------------------------------

    for fmt in [
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
    ]:

        try:

            parsed = datetime.strptime(
                text,
                fmt,
            )

            return parsed.strftime(
                "%m/%d/%Y %I:%M %p"
            )

        except Exception:

            continue


    # -------------------------------------------------
    # TRUE UTC / TIMEZONE-AWARE SOURCE
    # -------------------------------------------------

    try:

        parsed = pd.to_datetime(
            text,
            utc=True,
        )

        if pd.isna(
            parsed
        ):
            return text

        eastern = parsed.tz_convert(
            EASTERN
        )

        return eastern.strftime(
            "%m/%d/%Y %I:%M %p"
        )

    except Exception:

        return text


# =========================================================
# TABLE HTML
# =========================================================

def format_value(value):

    if pd.isna(value):

        return ""


    try:

        number = float(
            value
        )


        if number.is_integer():

            return (
                f"{int(number):,}"
            )


    except Exception:

        pass


    return html.escape(
        str(value)
    )


def dataframe_html(
    df,
    party="dem",
    max_rows=None,
):

    if df is None:

        return ""


    df = df.copy()

    if "Last Updated" in df.columns:

        df[
            "Last Updated"
        ] = (
            df[
                "Last Updated"
            ]
            .apply(
                format_eastern_timestamp
            )
        )

        df = df.rename(
            columns={
                "Last Updated":
                    "Last Vote Change"
            }
        )


    if max_rows is not None:

        df = df.head(
            max_rows
        )


    party_class = (
        "rep-table"
        if party == "rep"
        else ""
    )


    headers = "".join(
        f"<th>"
        f"{html.escape(str(column))}"
        f"</th>"

        for column in df.columns
    )


    rows = []


    for _, row in df.iterrows():

        cells = "".join(
            f"<td>"
            f"{format_value(row[column])}"
            f"</td>"

            for column in df.columns
        )


        rows.append(
            f"<tr>"
            f"{cells}"
            f"</tr>"
        )


    return f"""
<div class="results-table-wrap">

<table class="results-table {party_class}">

<thead>

<tr>
{headers}
</tr>

</thead>

<tbody>

{''.join(rows)}

</tbody>

</table>

</div>
"""


# =========================================================
# INITIAL FEDERAL DATA
# =========================================================

federal_dem = load_csv(
    FEDERAL_DEM_FILE
)

federal_rep = load_csv(
    FEDERAL_REP_FILE
)


# =========================================================
# STATUS
# =========================================================

status = load_status()


(
    federal_dem_units,
    federal_rep_units,
    federal_overall_units,
    federal_total_units,
) = get_reporting_status(
    status,
    dem=federal_dem,
    rep=federal_rep,
    prefix="federal",
)


last_checked = status.get(
    "last_checked",
    "Not available",
)

last_updated = status.get(
    "last_updated",
    "Not available",
)


# =========================================================
# HEADER
# =========================================================

logo_base64 = image_to_base64(
    LOGO
)


if logo_base64:

    logo_html = f"""
<img
src="data:image/png;base64,{logo_base64}"
class="ap-logo"
>
"""


else:

    logo_html = """
<div class="ap-logo-fallback">
AP
</div>
"""


@st.fragment(
    run_every="30s"
)
def render_header():

    header_status = load_status()

    election_status = (
        str(
            header_status.get(
                "election_status",
                "UNOFFICIAL",
            )
        )
        .strip()
        .upper()
    )

    if not election_status:

        election_status = (
            "UNOFFICIAL"
        )

    status_class = (
        "ap-unofficial"
        if election_status
        == "UNOFFICIAL"
        else "ap-result-status"
    )

    current_time = (
    datetime.now(
        NEW_YORK_TZ
    )
    .strftime(
        "%m/%d/%Y %I:%M %p"
    )
)

    st.html(
        f"""
<div class="top-header">

<div class="top-header-left">

{logo_html}

<div>

<div class="ap-header-title">
Vermont Election Results
</div>

<div class="ap-header-subtitle">
{ELECTION_NAME} • {ELECTION_DATE_DISPLAY}
</div>

</div>

</div>


<div class="ap-header-status">

<span class="{status_class}">
{html.escape(election_status)}
</span>

<span class="ap-pipe">
|
</span>

{current_time}

<span class="ap-pipe">
|
</span>

Next update:
00:30

</div>

</div>
"""
    )


render_header()


# =========================================================
# NAVIGATION
# =========================================================

def render_navigation():

    live_status = load_status()

    federal_dem_nav = load_csv(
        FEDERAL_DEM_FILE
    )

    federal_rep_nav = load_csv(
        FEDERAL_REP_FILE
    )

    (
        federal_dem_units,
        federal_rep_units,
        federal_overall_units,
        federal_total_units,
    ) = get_reporting_status(
        live_status,
        dem=federal_dem_nav,
        rep=federal_rep_nav,
        prefix="federal",
    )

    statewide_sidebar_units = int(
        live_status.get(
            "statewide_overall_reporting",
            0,
        )
    )

    statewide_sidebar_total = int(
        live_status.get(
            "statewide_total_units",
            DEFAULT_TOTAL_UNITS,
        )
    )

    statewide_sidebar = (
        f"{statewide_sidebar_units} / "
        f"{statewide_sidebar_total} RU"
        if statewide_sidebar_units > 0
        else "Not loaded"
    )

    last_checked = format_eastern_timestamp(
        live_status.get(
            "last_checked"
        )
    )

    last_updated = format_eastern_timestamp(
        live_status.get(
            "last_updated"
        )
    )

    federal_active = (
        "nav-active"
        if section == "federal"
        else ""
    )

    statewide_active = (
        "nav-active"
        if section == "statewide"
        else ""
    )

    st.html(
        f"""
<div class="nav-box">

<div class="nav-title">
Results
</div>

<a class="nav-link" href="?section=federal" target="_self">
<div class="nav-row {federal_active}">
<span class="dot-red">●</span>
Federal
</div>
</a>

<a class="nav-link" href="?section=statewide" target="_self">
<div class="nav-row {statewide_active}">
<span class="dot-white">●</span>
Statewide
</div>
</a>

<div class="nav-row">
<span class="dot-white">●</span>
Senate
</div>

<div class="nav-row">
<span class="dot-white">●</span>
House
</div>

<div class="nav-rule"></div>

<div class="nav-stat">
Federal:
<strong>
{federal_overall_units} / {federal_total_units} RU
</strong>
</div>

<div class="nav-stat">
Statewide:
<strong>
{statewide_sidebar}
</strong>
</div>

<div class="nav-stat">
Senate:
<strong>Not loaded</strong>
</div>

<div class="nav-stat">
House:
<strong>Not loaded</strong>
</div>

<div class="nav-rule" style="margin-top:150px;"></div>

<div class="nav-bottom-label">
Last checked
</div>

<div class="nav-bottom-value">
{html.escape(str(last_checked))}
</div>

<div class="nav-bottom-label">
Last results update
</div>

<div class="nav-bottom-value">
{html.escape(str(last_updated))}
</div>

<div class="nav-bottom-label">
Auto-refresh
</div>

<div class="nav-bottom-value">
every 30 seconds
<span class="green-dot">●</span>
</div>

</div>
"""
    )


# =========================================================
# FEDERAL PAGE
# =========================================================

def render_federal():

    dem = load_csv(
        FEDERAL_DEM_FILE
    )


    rep = load_csv(
        FEDERAL_REP_FILE
    )


    live_status = load_status()


    (
        dem_units,
        rep_units,
        overall_units,
        total_units,
    ) = get_reporting_status(
        live_status,
        dem=dem,
        rep=rep,
        prefix="federal",
    )


    live_last_checked = format_eastern_timestamp(
        live_status.get(
            "last_checked"
        )
    )


    live_last_updated = format_eastern_timestamp(
        live_status.get(
            "last_updated"
        )
    )


    st.html(
        f"""
<div class="main-title-row">

<div class="main-title">
Federal Results — {ELECTION_DATE_DISPLAY}
</div>

<div class="link-icon">
↗
</div>

</div>

<div class="main-subtitle">
{ELECTION_NAME}
</div>
"""
    )


    c1, c2, c3 = st.columns(
        3
    )


    with c1:

        st.html(
            f"""
<div class="metric-card">

<div class="metric-inner">

<div class="
status-circle
status-circle-dem
">
✓
</div>

<div>

<div class="metric-label">
Federal Democratic
</div>

<div class="
metric-value
metric-dem
">

{
    "Loaded"
    if dem is not None
    else "Waiting"
}

</div>

<div class="metric-sub">

Reporting Units:
{dem_units} / {total_units}

</div>

</div>

</div>

</div>
"""
        )


    with c2:

        st.html(
            f"""
<div class="metric-card">

<div class="metric-inner">

<div class="
status-circle
status-circle-rep
">
✓
</div>

<div>

<div class="metric-label">
Federal Republican
</div>

<div class="
metric-value
metric-rep
">

{
    "Loaded"
    if rep is not None
    else "Waiting"
}

</div>

<div class="metric-sub">

Reporting Units:
{rep_units} / {total_units}

</div>

</div>

</div>

</div>
"""
        )


    with c3:

        st.html(
            f"""
<div class="metric-card">

<div class="metric-inner">

<div class="
status-circle
status-circle-neutral
">
♙
</div>

<div>

<div class="metric-label">
Federal Reporting Units
</div>

<div class="metric-value">

{total_units}

</div>

<div class="metric-sub">
Total Reporting Units
</div>

</div>

</div>

</div>
"""
        )


    st.html(
        f"""
<div class="status-line">

Last checked:
{html.escape(str(live_last_checked))}

&nbsp;•&nbsp;

Last results update:
{html.escape(str(live_last_updated))}

&nbsp;•&nbsp;

Auto-refresh:
30 seconds

</div>
"""
    )


    dem_tab, rep_tab = st.tabs(
        [
            "Democratic",
            "Republican",
        ]
    )


    # =====================================================
    # FEDERAL DEM
    # =====================================================

    with dem_tab:

        st.html(
            f"""
<div class="result-card">

<div class="
result-heading
result-heading-dem
">

Democratic — Federal Results

</div>

<div class="result-ru">

{ELECTION_DATE_DISPLAY}

&nbsp; • &nbsp;

Reporting Units:
{dem_units} / {total_units}

</div>

</div>
"""
        )


        search = st.text_input(
            "Search Federal Democratic",
            placeholder=
                "Search town or candidate...",
            label_visibility=
                "collapsed",
            key=
                "federal_dem_search",
        )


        if dem is not None:

            view = filter_results(
                dem,
                search,
            )


            view = clean_results(
                view,
                keep_districts=True,
            )


            rows_to_show = (
                None

                if st.session_state[
                    "federal_dem_full"
                ]

                else 5
            )


            st.html(
                dataframe_html(
                    view,
                    party="dem",
                    max_rows=
                        rows_to_show,
                )
            )


            label = (
                "SHOW FEWER DEMOCRATIC RESULTS"

                if st.session_state[
                    "federal_dem_full"
                ]

                else

                "VIEW FULL DEMOCRATIC RESULTS"
            )


            if st.button(
                label,
                key=
                    "federal_dem_full_button",
            ):

                st.session_state[
                    "federal_dem_full"
                ] = (
                    not
                    st.session_state[
                        "federal_dem_full"
                    ]
                )


                st.rerun()


        else:

            st.warning(
                "Waiting for Democratic "
                "Federal results."
            )


    # =====================================================
    # FEDERAL REP
    # =====================================================

    with rep_tab:

        st.html(
            f"""
<div class="result-card">

<div class="
result-heading
result-heading-rep
">

Republican — Federal Results

</div>

<div class="result-ru">

{ELECTION_DATE_DISPLAY}

&nbsp; • &nbsp;

Reporting Units:
{rep_units} / {total_units}

</div>

</div>
"""
        )


        search = st.text_input(
            "Search Federal Republican",
            placeholder=
                "Search town or candidate...",
            label_visibility=
                "collapsed",
            key=
                "federal_rep_search",
        )


        if rep is not None:

            view = filter_results(
                rep,
                search,
            )


            view = clean_results(
                view,
                keep_districts=True,
            )


            rows_to_show = (
                None

                if st.session_state[
                    "federal_rep_full"
                ]

                else 5
            )


            st.html(
                dataframe_html(
                    view,
                    party="rep",
                    max_rows=
                        rows_to_show,
                )
            )


            label = (
                "SHOW FEWER REPUBLICAN RESULTS"

                if st.session_state[
                    "federal_rep_full"
                ]

                else

                "VIEW FULL REPUBLICAN RESULTS"
            )


            if st.button(
                label,
                key=
                    "federal_rep_full_button",
            ):

                st.session_state[
                    "federal_rep_full"
                ] = (
                    not
                    st.session_state[
                        "federal_rep_full"
                    ]
                )


                st.rerun()


        else:

            st.warning(
                "Waiting for Republican "
                "Federal results."
            )


# =========================================================
# STATEWIDE PAGE
# =========================================================

def render_statewide():

    live_status = load_status()


    st.html(
        f"""
<div class="main-title-row">

<div class="main-title">
Statewide Results — {ELECTION_DATE_DISPLAY}
</div>

<div class="link-icon">
↗
</div>

</div>

<div class="main-subtitle">
{ELECTION_NAME}
</div>
"""
    )


    st.html(
        """
<div class="office-label">
SELECT OFFICE
</div>
"""
    )


    selected_office = st.selectbox(
        "Select Office",
        STATEWIDE_OFFICES,
        label_visibility="collapsed",
        key="statewide_office",
    )


    office_files = (
        STATEWIDE_FILES[
            selected_office
        ]
    )


    dem_file = (
        STATEWIDE_DEM_DIR
        /
        office_files[
            "dem"
        ]
    )


    rep_file = (
        STATEWIDE_REP_DIR
        /
        office_files[
            "rep"
        ]
    )


    dem = load_excel(
        dem_file
    )


    rep = load_excel(
        rep_file
    )


    (
        dem_units,
        rep_units,
        overall_units,
        total_units,
    ) = get_reporting_status(
        live_status,
        dem=dem,
        rep=rep,
        prefix="statewide",
    )


    live_last_checked = format_eastern_timestamp(
        live_status.get(
            "last_checked"
        )
    )


    live_last_updated = format_eastern_timestamp(
        live_status.get(
            "last_updated"
        )
    )


    # =====================================================
    # CARDS
    # =====================================================

    c1, c2, c3 = st.columns(
        3
    )


    with c1:

        st.html(
            f"""
<div class="metric-card">

<div class="metric-inner">

<div class="
status-circle
status-circle-dem
">
✓
</div>

<div>

<div class="metric-label">
Statewide Democratic
</div>

<div class="
metric-value
metric-dem
">

{
    "Loaded"
    if dem is not None
    else "Waiting"
}

</div>

<div class="metric-sub">

Reporting Units:
{dem_units} / {total_units}

</div>

</div>

</div>

</div>
"""
        )


    with c2:

        st.html(
            f"""
<div class="metric-card">

<div class="metric-inner">

<div class="
status-circle
status-circle-rep
">
✓
</div>

<div>

<div class="metric-label">
Statewide Republican
</div>

<div class="
metric-value
metric-rep
">

{
    "Loaded"
    if rep is not None
    else "Waiting"
}

</div>

<div class="metric-sub">

Reporting Units:
{rep_units} / {total_units}

</div>

</div>

</div>

</div>
"""
        )


    with c3:

        st.html(
            f"""
<div class="metric-card">

<div class="metric-inner">

<div class="
status-circle
status-circle-neutral
">
♙
</div>

<div>

<div class="metric-label">
Statewide Reporting Units
</div>

<div class="metric-value">

{total_units}

</div>

<div class="metric-sub">
Total Reporting Units
</div>

</div>

</div>

</div>
"""
        )


    st.html(
        f"""
<div class="status-line">

Last checked:
{html.escape(str(live_last_checked))}

&nbsp;•&nbsp;

Last results update:
{html.escape(str(live_last_updated))}

&nbsp;•&nbsp;

Auto-refresh:
30 seconds

</div>
"""
    )


    dem_tab, rep_tab = st.tabs(
        [
            "Democratic",
            "Republican",
        ]
    )


    pretty_office = (
        selected_office
        .title()
    )


    # =====================================================
    # STATEWIDE DEM
    # =====================================================

    with dem_tab:

        st.html(
            f"""
<div class="result-card">

<div class="
result-heading
result-heading-dem
">

{pretty_office} — Democratic

</div>

<div class="result-ru">

{ELECTION_DATE_DISPLAY}

&nbsp; • &nbsp;

Reporting Units:
{dem_units} / {total_units}

</div>

</div>
"""
        )


        search = st.text_input(
            "Search Statewide Democratic",
            placeholder=
                "Search town or candidate...",
            label_visibility=
                "collapsed",
            key=
                "statewide_dem_search",
        )


        if dem is not None:

            view = filter_results(
                dem,
                search,
            )


            view = clean_results(
                view,
                keep_districts=True,
            )


            rows_to_show = (
                None

                if st.session_state[
                    "statewide_dem_full"
                ]

                else 5
            )


            st.html(
                dataframe_html(
                    view,
                    party="dem",
                    max_rows=
                        rows_to_show,
                )
            )


            label = (
                "SHOW FEWER DEMOCRATIC RESULTS"

                if st.session_state[
                    "statewide_dem_full"
                ]

                else

                "VIEW FULL DEMOCRATIC RESULTS"
            )


            if st.button(
                label,
                key=
                    "statewide_dem_full_button",
            ):

                st.session_state[
                    "statewide_dem_full"
                ] = (
                    not
                    st.session_state[
                        "statewide_dem_full"
                    ]
                )


                st.rerun()


        else:

            st.warning(
                f"Waiting for Democratic "
                f"{pretty_office} results."
            )


    # =====================================================
    # STATEWIDE REP
    # =====================================================

    with rep_tab:

        st.html(
            f"""
<div class="result-card">

<div class="
result-heading
result-heading-rep
">

{pretty_office} — Republican

</div>

<div class="result-ru">

{ELECTION_DATE_DISPLAY}

&nbsp; • &nbsp;

Reporting Units:
{rep_units} / {total_units}

</div>

</div>
"""
        )


        search = st.text_input(
            "Search Statewide Republican",
            placeholder=
                "Search town or candidate...",
            label_visibility=
                "collapsed",
            key=
                "statewide_rep_search",
        )


        if rep is not None:

            view = filter_results(
                rep,
                search,
            )


            view = clean_results(
                view,
                keep_districts=True,
            )


            rows_to_show = (
                None

                if st.session_state[
                    "statewide_rep_full"
                ]

                else 5
            )


            st.html(
                dataframe_html(
                    view,
                    party="rep",
                    max_rows=
                        rows_to_show,
                )
            )


            label = (
                "SHOW FEWER REPUBLICAN RESULTS"

                if st.session_state[
                    "statewide_rep_full"
                ]

                else

                "VIEW FULL REPUBLICAN RESULTS"
            )


            if st.button(
                label,
                key=
                    "statewide_rep_full_button",
            ):

                st.session_state[
                    "statewide_rep_full"
                ] = (
                    not
                    st.session_state[
                        "statewide_rep_full"
                    ]
                )


                st.rerun()


        else:

            st.warning(
                f"Waiting for Republican "
                f"{pretty_office} results."
            )


# =========================================================
# BACKGROUND REFRESH
# =========================================================

@st.cache_resource
def get_refresh_state():

    return {
        "lock":
            threading.Lock(),

        "running":
            False,

        "last_started":
            0.0,

        "error":
            None,
    }


def run_refresh_in_background(
    state,
):

    try:

        asyncio.run(
            check_results()
        )

        error = None

    except Exception as exc:

        error = str(
            exc
        )

    finally:

        with state[
            "lock"
        ]:

            state[
                "running"
            ] = False

            state[
                "error"
            ] = error


def trigger_background_refresh():

    state = get_refresh_state()

    now = time.monotonic()

    with state[
        "lock"
    ]:

        # Do not start a second request while one is still
        # running, and do not start more often than every
        # ~25 seconds even if Streamlit reruns for another
        # reason such as a button click.

        if state[
            "running"
        ]:

            return

        if (
            now
            - state[
                "last_started"
            ]
            < 25
        ):

            return

        state[
            "running"
        ] = True

        state[
            "last_started"
        ] = now

    thread = threading.Thread(
        target=
            run_refresh_in_background,

        args=(
            state,
        ),

        daemon=True,
    )

    thread.start()


def get_background_refresh_error():

    state = get_refresh_state()

    with state[
        "lock"
    ]:

        return state.get(
            "error"
        )


# =========================================================
# LIVE DASHBOARD
# =========================================================

@st.fragment(
    run_every="30s"
)
def live_dashboard():

    # Start the network/file refresh in the background.
    # The current UI remains visible while it runs.
    trigger_background_refresh()

    refresh_error = (
        get_background_refresh_error()
    )

    nav_col, main_col = st.columns(
        [
            1.3,
            8.7,
        ],
        gap="small",
    )

    with nav_col:

        render_navigation()

    with main_col:

        if refresh_error:

            st.warning(
                f"Could not refresh results: "
                f"{refresh_error}"
            )

        if section == "statewide":

            render_statewide()

        else:

            render_federal()


live_dashboard()