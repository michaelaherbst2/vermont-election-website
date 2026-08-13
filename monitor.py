# To run:
# uv run python monitor.py

from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import asyncio
import hashlib
import json
import os
import tempfile

import httpx
import pandas as pd


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

STATUS_FILE = RESULTS_DIR / "status.json"

RU_TIMESTAMPS_FILE = (
    RESULTS_DIR / "ru_timestamps.json"
)

STATEWIDE_RAW_JSON = (
    STATEWIDE_DIR / "statewide_live.json"
)

CHECK_INTERVAL = 30

RU_TIMESTAMP_HISTORY_VERSION = 6


# =========================================================
# FEDERAL FILES
# =========================================================

FEDERAL_DEM_FILE = (
    FEDERAL_DIR / "democratic_results.csv"
)

FEDERAL_REP_FILE = (
    FEDERAL_DIR / "republican_results.csv"
)

# One-time migration source for existing Federal files.
# If the CSV does not exist yet but the old XLSX does,
# monitor.py will convert the XLSX to CSV automatically.
LEGACY_FEDERAL_DEM_FILE = (
    FEDERAL_DIR / "democratic_results.xlsx"
)

LEGACY_FEDERAL_REP_FILE = (
    FEDERAL_DIR / "republican_results.xlsx"
)


# =========================================================
# LIVE VERMONT ELECTION
# =========================================================

ELECTION_GUID = (
    "a18f77e0-89f8-4a01-8d97-61a7c75ba200"
)

STATIC_BASE_URL = (
    "https://static.electionresults.vermont.gov/"
)

MANIFEST_URL = (
    STATIC_BASE_URL
    + "elections/"
    + ELECTION_GUID
    + ".json"
)


# =========================================================
# STATEWIDE FILE MAP
# =========================================================

STATEWIDE_FILES = {
    "GOVERNOR": {
        "dem":
            STATEWIDE_DEM_DIR
            / "dem_governor.xlsx",

        "rep":
            STATEWIDE_REP_DIR
            / "rep_governor.xlsx",
    },

    "LIEUTENANT GOVERNOR": {
        "dem":
            STATEWIDE_DEM_DIR
            / "dem_lieutenant_governor.xlsx",

        "rep":
            STATEWIDE_REP_DIR
            / "rep_lieutenant_governor.xlsx",
    },

    "STATE TREASURER": {
        "dem":
            STATEWIDE_DEM_DIR
            / "dem_state_treasurer.xlsx",

        "rep":
            STATEWIDE_REP_DIR
            / "rep_state_treasurer.xlsx",
    },

    "SECRETARY OF STATE": {
        "dem":
            STATEWIDE_DEM_DIR
            / "dem_secretary_of_state.xlsx",

        "rep":
            STATEWIDE_REP_DIR
            / "rep_secretary_of_state.xlsx",
    },

    "AUDITOR OF ACCOUNTS": {
        "dem":
            STATEWIDE_DEM_DIR
            / "dem_auditor_of_accounts.xlsx",

        "rep":
            STATEWIDE_REP_DIR
            / "rep_auditor_of_accounts.xlsx",
    },

    "ATTORNEY GENERAL": {
        "dem":
            STATEWIDE_DEM_DIR
            / "dem_attorney_general.xlsx",

        "rep":
            STATEWIDE_REP_DIR
            / "rep_attorney_general.xlsx",
    },
}


# =========================================================
# CREATE DIRECTORIES
# =========================================================

for directory in [
    RESULTS_DIR,
    FEDERAL_DIR,
    STATEWIDE_DIR,
    STATEWIDE_DEM_DIR,
    STATEWIDE_REP_DIR,
]:

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )


# =========================================================
# TIME
# =========================================================

def now_string():

    return datetime.now(
        NEW_YORK_TZ
    ).strftime(
        "%m/%d/%Y %I:%M:%S %p"
    )



# =========================================================
# SAFE INTEGER
# =========================================================

def safe_int(
    value,
    default=0,
):

    try:

        if value is None:
            return default

        return int(
            float(value)
        )

    except Exception:

        return default



# =========================================================
# ATOMIC FILE WRITES
# =========================================================

def atomic_write_text(
    path,
    text,
):

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        dir=path.parent,
        suffix=".tmp",
        delete=False,
    ) as temp_file:

        temp_path = Path(
            temp_file.name
        )

        temp_file.write(
            text
        )

    os.replace(
        temp_path,
        path,
    )


def atomic_to_csv(
    df,
    path,
):

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        newline="",
        dir=path.parent,
        suffix=".csv",
        delete=False,
    ) as temp_file:

        temp_path = Path(
            temp_file.name
        )

    try:

        df.to_csv(
            temp_path,
            index=False,
        )

        os.replace(
            temp_path,
            path,
        )

    finally:

        if temp_path.exists():

            try:
                temp_path.unlink()
            except Exception:
                pass


def atomic_to_excel(
    df,
    path,
):

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with tempfile.NamedTemporaryFile(
        dir=path.parent,
        suffix=".xlsx",
        delete=False,
    ) as temp_file:

        temp_path = Path(
            temp_file.name
        )

    try:

        df.to_excel(
            temp_path,
            index=False,
        )

        os.replace(
            temp_path,
            path,
        )

    finally:

        if temp_path.exists():

            try:
                temp_path.unlink()
            except Exception:
                pass


# =========================================================
# STATUS.JSON
# =========================================================

def load_status():

    if not STATUS_FILE.exists():
        return {}

    try:

        return json.loads(
            STATUS_FILE.read_text()
        )

    except Exception:

        return {}


def save_status(**updates):

    status = load_status()

    for key, value in updates.items():

        if value is not None:
            status[key] = value

    if (
        "error" in updates
        and updates["error"] is None
    ):

        status.pop(
            "error",
            None,
        )

    atomic_write_text(
        STATUS_FILE,
        json.dumps(
            status,
            indent=2,
        ),
    )


# =========================================================
# RU TIMESTAMP HISTORY
# =========================================================

def load_ru_timestamps():

    """
    Load persistent per-town vote-change history.

    History is preserved across app restarts and code updates.
    """

    if not RU_TIMESTAMPS_FILE.exists():

        return {
            "_version":
                RU_TIMESTAMP_HISTORY_VERSION,
        }

    try:

        data = json.loads(
            RU_TIMESTAMPS_FILE.read_text()
        )

    except Exception:

        return {
            "_version":
                RU_TIMESTAMP_HISTORY_VERSION,
        }

    if not isinstance(
        data,
        dict,
    ):

        return {
            "_version":
                RU_TIMESTAMP_HISTORY_VERSION,
        }

    data[
        "_version"
    ] = RU_TIMESTAMP_HISTORY_VERSION

    return data



def save_ru_timestamps(data):

    data[
        "_version"
    ] = RU_TIMESTAMP_HISTORY_VERSION

    atomic_write_text(
        RU_TIMESTAMPS_FILE,
        json.dumps(
            data,
            indent=2,
        ),
    )


# =========================================================
# ROW SIGNATURE
# =========================================================

def row_signature(row):

    """
    JSON-stable signature of vote/count values only.

    A new Last Vote Change is recorded only when a numeric
    result value changes for that town.
    """

    ignored_fields = {
        "Last Updated",
        "Town",
        "Rep District",
        "Sen District",
        "_source_last_updated",
        "lastUpdated",
        "lastUpdatedDate",
        "updated",
        "updateTime",
        "timestamp",
    }

    signature = []

    for key, value in row.items():

        if key in ignored_fields:
            continue

        if pd.isna(value):
            continue

        try:

            number = float(
                str(value)
                .strip()
                .replace(
                    ",",
                    "",
                )
            )

        except Exception:
            continue

        signature.append(
            [
                str(key),
                number,
            ]
        )

    return signature


def row_is_reporting(row):

    """
    Prefer the source Total when available.

    This prevents an entirely blank/non-reporting town from
    receiving a misleading update timestamp.
    """

    for total_name in [
        "Total",
        "Total Votes",
    ]:

        if total_name in row:

            try:

                return (
                    float(
                        row[
                            total_name
                        ]
                    )
                    > 0
                )

            except Exception:

                pass

    ignored = {
        "Town",
        "Rep District",
        "Sen District",
        "Last Updated",
    }

    for key, value in row.items():

        if key in ignored:
            continue

        try:

            if float(value) > 0:
                return True

        except Exception:
            continue

    return False


# =========================================================
# SOURCE TIME TO EASTERN
# =========================================================

def source_time_to_eastern(value):

    if value is None:
        return ""

    text = str(
        value
    ).strip()

    if not text:
        return ""

    try:

        parsed = datetime.fromisoformat(
            text.replace(
                "Z",
                "+00:00",
            )
        )

        # If Vermont provides an actual timezone,
        # convert it to New York time.
        if parsed.tzinfo is not None:

            parsed = parsed.astimezone(
                NEW_YORK_TZ
            )

        return parsed.strftime(
            "%m/%d/%Y %I:%M:%S %p"
        )

    except Exception:

        # Vermont currently gives some timestamps in
        # display format such as:
        #
        # 08/13/2026 10:26 AM
        #
        # Leave those unchanged.
        return text


# =========================================================
# SOURCE RU TIMESTAMP
# =========================================================

def get_source_ru_timestamp(row):

    """
    Use a reporting-unit timestamp supplied by Vermont when
    one is actually present in the parsed row.

    If Vermont does not provide one, return blank and the
    monitor will use the time it first observes a vote change.
    """

    possible_fields = [
        "_source_last_updated",
        "lastUpdated",
        "lastUpdatedDate",
        "updated",
        "updateTime",
        "timestamp",
    ]

    for field in possible_fields:

        value = row.get(
            field
        )

        if value is None:
            continue

        text = str(
            value
        ).strip()

        if (
            text
            and text.casefold() != "nan"
        ):
            return source_time_to_eastern(
                text
            )

    return ""


# =========================================================
# ADD / UPDATE RU TIMESTAMPS
# =========================================================

def add_ru_timestamps(
    df,
    category,
    party,
    office,
    observed_at=None,
    baseline_at=None,
):

    if (
        df is None
        or df.empty
        or "Town" not in df.columns
    ):

        return df


    output = df.copy()

    history = load_ru_timestamps()

    history_key = (
        f"{category}"
        f"|{party}"
        f"|{office}"
    )

    section_history = history.setdefault(
        history_key,
        {}
    )

    if observed_at is None:

        observed_at = now_string()

    if baseline_at is None:

        baseline_at = observed_at

    timestamps = []


    for _, row in output.iterrows():

        town = (
            str(
                row.get(
                    "Town",
                    "",
                )
            )
            .strip()
            .upper()
        )

        if not town:

            timestamps.append(
                ""
            )

            continue


        signature = row_signature(
            row.to_dict()
        )

        reporting = row_is_reporting(
            row.to_dict()
        )

        source_changed_at = (
            get_source_ru_timestamp(
                row.to_dict()
            )
        )

        previous = section_history.get(
            town
        )


        # -------------------------------------------------
        # FIRST OBSERVATION = GIVE EVERY REPORTING TOWN
        # AN INITIAL SOURCE TIMESTAMP
        # -------------------------------------------------

        if previous is None:

            # First observation = baseline timestamp.
            # Vermont does not provide a historical per-town
            # change time, so use the overall source update time
            # (or observed time) once.
            section_history[
                town
            ] = {
                "signature":
                    signature,

                "last_updated":
                    (
                        baseline_at
                        or observed_at
                    ),

                "reporting":
                    reporting,
            }


        # -------------------------------------------------
        # ACTUAL RESULT CHANGE OBSERVED
        # -------------------------------------------------

        elif (
            previous.get(
                "signature"
            )
            != signature
        ):

            # A real numeric vote/count difference was observed.
            # Only this town gets a new timestamp.

            previous_reporting = bool(
                previous.get(
                    "reporting",
                    False,
                )
            )

            if (
                reporting
                or previous_reporting
            ):

                # This is the exact poll when OUR monitor first
                # observed this town's vote/count values change.
                changed_at = observed_at

            else:

                changed_at = (
                    previous.get(
                        "last_updated",
                        "",
                    )
                )

            section_history[
                town
            ] = {
                "signature":
                    signature,

                "last_updated":
                    changed_at,

                "reporting":
                    reporting,
            }

            print(
                "VOTE CHANGE:",
                category,
                party,
                office,
                town,
                "->",
                changed_at,
            )


        # -------------------------------------------------
        # NO CHANGE = KEEP PREVIOUS TIME
        # -------------------------------------------------

        else:

            # No vote/result change:
            # preserve this town's existing Last Vote Change.
            section_history[
                town
            ][
                "reporting"
            ] = reporting

            # Only backfill genuinely blank legacy rows once.
            if not section_history[
                town
            ].get(
                "last_updated"
            ):

                section_history[
                    town
                ][
                    "last_updated"
                ] = (
                    source_changed_at
                    or baseline_at
                    or observed_at
                )


        timestamps.append(
            section_history[
                town
            ].get(
                "last_updated",
                "",
            )
        )


    history[
        history_key
    ] = section_history

    save_ru_timestamps(
        history
    )

    output[
        "Last Updated"
    ] = timestamps

    return output


# =========================================================
# COMBINE UNIQUE TEXT
# =========================================================

def combine_unique(values):

    result = []

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

        if text not in result:

            result.append(
                text
            )

    return ", ".join(
        result
    )


# =========================================================
# PARTY
# =========================================================

def get_party_key(
    party_name,
):

    party_name = (
        str(
            party_name
            or ""
        )
        .strip()
        .upper()
    )

    if "DEMOCRAT" in party_name:
        return "dem"

    if "REPUBLICAN" in party_name:
        return "rep"

    return None


# =========================================================
# ELECTION STATUS
# =========================================================

def get_election_status(manifest):

    """
    Return the election/results status exactly as Vermont
    exposes it when possible.

    We do not assume the final wording will be CERTIFIED.
    The feed might use OFFICIAL, FINAL, CERTIFIED, etc.

    If none of the known status fields are present, fall
    back to UNOFFICIAL.
    """

    likely_keys = {
        "status",
        "resultStatus",
        "resultsStatus",
        "electionStatus",
        "certificationStatus",
        "statusText",
        "resultStatusText",
        "resultsStatusText",
    }

    # First check the top level.
    for key in likely_keys:

        value = manifest.get(
            key
        )

        if value is None:
            continue

        text = str(
            value
        ).strip()

        if text:
            return text.upper()

    # Then search nested dictionaries/lists for a likely
    # status key. This makes the code more tolerant if the
    # Vermont manifest moves the field later.
    def search_nested(value):

        if isinstance(
            value,
            dict,
        ):

            for key, child in value.items():

                if key in likely_keys:

                    text = str(
                        child
                    ).strip()

                    if text:
                        return text.upper()

            for child in value.values():

                found = search_nested(
                    child
                )

                if found:
                    return found

        elif isinstance(
            value,
            list,
        ):

            for child in value:

                found = search_nested(
                    child
                )

                if found:
                    return found

        return None

    found = search_nested(
        manifest
    )

    if found:
        return found

    return "UNOFFICIAL"


# =========================================================
# URL
# =========================================================

def build_static_url(path):

    if not path:
        return None

    clean_path = (
        str(path)
        .replace(
            "\\",
            "/",
        )
        .lstrip("/")
    )

    return (
        STATIC_BASE_URL
        + clean_path
    )


# =========================================================
# FETCH JSON
# =========================================================

async def fetch_json(
    client,
    url,
):

    response = await client.get(
        url
    )

    response.raise_for_status()

    text = (
        response.content
        .decode(
            "utf-8-sig"
        )
    )

    return json.loads(
        text
    )


# =========================================================
# BUILD ONE STATEWIDE OFFICE
# =========================================================

def build_statewide_office(
    office_block,
):

    town_blocks = (
        office_block.get(
            "cs",
            []
        )
        or []
    )


    rows = []


    for town_block in town_blocks:

        town = (
            str(
                town_block.get(
                    "tn",
                    ""
                )
                or ""
            )
            .strip()
            .upper()
        )


        if not town:
            continue


        # Skip statewide summary / aggregate rows.
        # Only actual town reporting units belong in the
        # per-town results table.
        if is_summary_reporting_unit(
            town
        ):
            continue


        row = {
            "Town":
                town,

            "_source_last_updated":
                (
                    town_block.get("lastUpdated")
                    or town_block.get("lastUpdatedDate")
                    or town_block.get("updated")
                    or town_block.get("updateTime")
                    or town_block.get("timestamp")
                    or ""
                ),

            "Rep District":
                str(
                    town_block.get(
                        "repd",
                        ""
                    )
                    or ""
                )
                .strip(),

            "Sen District":
                str(
                    town_block.get(
                        "send",
                        ""
                    )
                    or ""
                )
                .strip(),
        }


        # =================================================
        # REGULAR CANDIDATES
        # =================================================

        regular_candidates = (
            town_block.get(
                "rc",
                []
            )
            or []
        )


        for candidate in regular_candidates:

            candidate_name = (
                str(
                    candidate.get(
                        "cn",
                        ""
                    )
                    or ""
                )
                .strip()
                .upper()
            )


            if not candidate_name:
                continue


            votes = safe_int(
                candidate.get(
                    "vc"
                )
            )


            row[
                candidate_name
            ] = (
                row.get(
                    candidate_name,
                    0,
                )
                + votes
            )


        # =================================================
        # WRITE-INS
        # =================================================

        total_write_ins = 0


        write_ins = (
            town_block.get(
                "wc",
                []
            )
            or []
        )


        for write_in in write_ins:

            total_write_ins += (
                safe_int(
                    write_in.get(
                        "vc"
                    )
                )
            )


        row[
            "Total Write Ins"
        ] = total_write_ins


        # =================================================
        # OTHERS
        # =================================================

        others = None


        for key in [
            "others",
            "Others",
            "oc",
        ]:

            if key in town_block:

                others = safe_int(
                    town_block.get(
                        key
                    )
                )

                break


        if others is not None:

            row[
                "Others"
            ] = others


        # =================================================
        # SOURCE TOTAL
        # =================================================

        if "sc" in town_block:

            row[
                "Total"
            ] = safe_int(
                town_block.get(
                    "sc"
                )
            )


        rows.append(
            row
        )


    if not rows:

        return pd.DataFrame()


    df = pd.DataFrame(
        rows
    )


    # =====================================================
    # CLEAN TYPES
    # =====================================================

    text_columns = {
        "Town",
        "Rep District",
        "Sen District",
        "_source_last_updated",
    }


    for column in df.columns:

        if column in text_columns:

            df[
                column
            ] = (
                df[
                    column
                ]
                .fillna("")
                .astype(str)
            )

            continue


        df[
            column
        ] = (
            pd.to_numeric(
                df[
                    column
                ],
                errors="coerce",
            )
            .fillna(0)
            .astype(int)
        )


    # =====================================================
    # COMBINE DUPLICATE TOWNS
    # =====================================================

    aggregation = {}


    for column in df.columns:

        if column == "Town":
            continue


        if column == "_source_last_updated":

            aggregation[
                column
            ] = lambda values: max(
                [
                    str(value).strip()
                    for value in values
                    if (
                        not pd.isna(value)
                        and str(value).strip()
                    )
                ],
                default="",
            )

        elif column in {
            "Rep District",
            "Sen District",
        }:

            aggregation[
                column
            ] = combine_unique


        else:

            aggregation[
                column
            ] = "sum"


    df = (
        df.groupby(
            "Town",
            as_index=False,
        )
        .agg(
            aggregation
        )
    )


    # =====================================================
    # REMOVE ZERO-ONLY CANDIDATES
    # =====================================================

    protected = {
        "Town",
        "Rep District",
        "Sen District",
        "Total Write Ins",
        "Others",
        "Total",
    }


    drop_columns = []


    for column in df.columns:

        if column in protected:
            continue


        numeric = pd.to_numeric(
            df[
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

        df = df.drop(
            columns=
                drop_columns
        )


    # =====================================================
    # COLUMN ORDER
    # =====================================================

    first_columns = [
        column
        for column in [
            "Town",
            "Rep District",
            "Sen District",
        ]
        if column in df.columns
    ]


    near_end = [
        column
        for column in [
            "Total Write Ins",
            "Others",
        ]
        if column in df.columns
    ]


    candidate_columns = [
        column
        for column in df.columns
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


    if "Total" in df.columns:

        ordered.append(
            "Total"
        )


    df = df[
        ordered
    ]

    if "_source_last_updated" in df.columns:

        # Keep this internal metadata only until timestamps
        # have been assigned; it is not a display column.
        pass


    return (
        df
        .sort_values(
            "Town"
        )
        .reset_index(
            drop=True
        )
    )


# =========================================================
# SORT BY LAST UPDATED
# =========================================================

def sort_by_last_updated(df):

    if (
        df is None
        or df.empty
        or "Last Updated" not in df.columns
    ):

        return df


    output = df.copy()


    output[
        "_timestamp_sort"
    ] = pd.to_datetime(
        output[
            "Last Updated"
        ],
        errors="coerce",
    )


    output = (
        output
        .sort_values(
            [
                "_timestamp_sort",
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
                "_timestamp_sort"
            ]
        )
        .reset_index(
            drop=True
        )
    )


    return output


# =========================================================
# PUT LAST UPDATED AT FRONT
# =========================================================

def order_timestamp_column(df):

    if (
        df is None
        or "Last Updated" not in df.columns
    ):

        return df


    first = []


    for column in [
        "Last Updated",
        "Town",
        "Rep District",
        "Sen District",
    ]:

        if column in df.columns:

            first.append(
                column
            )


    remaining = [
        column
        for column in df.columns
        if column not in first
    ]


    return df[
        first
        +
        remaining
    ]



# =========================================================
# SUMMARY REPORTING UNITS
# =========================================================

def is_summary_reporting_unit(name):

    normalized = (
        str(
            name
            or ""
        )
        .strip()
        .upper()
        .replace("-", " ")
        .replace("_", " ")
    )

    normalized = " ".join(
        normalized.split()
    )

    return normalized in {
        "STATE",
        "STATE WIDE",
        "STATEWIDE",
        "STATE TOTAL",
        "STATE TOTALS",
        "TOTAL",
        "TOTALS",
    }


# =========================================================
# FEDERAL HELPERS
# =========================================================

FEDERAL_OFFICE_NAMES = {
    "REPRESENTATIVE TO CONGRESS",
    "US REPRESENTATIVE",
    "U.S. REPRESENTATIVE",
    "UNITED STATES REPRESENTATIVE",
    "REPRESENTATIVE",
}


def normalize_office_name(value):

    return (
        str(
            value
            or ""
        )
        .strip()
        .upper()
    )


def build_federal_office(
    office_block,
):

    """
    Build one Federal office into separate Democratic and
    Republican town DataFrames.

    Vermont's live result payloads use the same compact
    candidate/town keys seen in the Statewide feed:
      tn   = town
      repd = representative district
      send = senate district
      rc   = regular candidates
      wc   = write-ins
      sc   = source total

    Candidate party may appear on the candidate row. If not,
    the surrounding party block supplies the party.
    """

    town_blocks = (
        office_block.get(
            "cs",
            []
        )
        or office_block.get(
            "towns",
            []
        )
        or []
    )

    rows_by_party = {
        "dem": [],
        "rep": [],
    }

    for town_block in town_blocks:

        town = (
            str(
                town_block.get(
                    "tn",
                    town_block.get(
                        "town",
                        "",
                    ),
                )
                or ""
            )
            .strip()
            .upper()
        )

        if not town:
            continue

        # Skip statewide/total aggregate rows from the
        # Federal feed. Only actual town reporting units
        # belong in the per-town Federal table.
        if is_summary_reporting_unit(
            town
        ):
            continue

        base = {
            "Town":
                town,

            "_source_last_updated":
                (
                    town_block.get("lastUpdated")
                    or town_block.get("lastUpdatedDate")
                    or town_block.get("updated")
                    or town_block.get("updateTime")
                    or town_block.get("timestamp")
                    or ""
                ),

            "Rep District":
                str(
                    town_block.get(
                        "repd",
                        town_block.get(
                            "repDistrict",
                            "",
                        ),
                    )
                    or ""
                )
                .strip(),

            "Sen District":
                str(
                    town_block.get(
                        "send",
                        town_block.get(
                            "senDistrict",
                            "",
                        ),
                    )
                    or ""
                )
                .strip(),
        }

        party_rows = {
            "dem": dict(base),
            "rep": dict(base),
        }

        party_has_candidate = {
            "dem": False,
            "rep": False,
        }

        regular_candidates = (
            town_block.get(
                "rc",
                []
            )
            or town_block.get(
                "candidates",
                []
            )
            or []
        )

        for candidate in regular_candidates:

            candidate_name = (
                str(
                    candidate.get(
                        "cn",
                        candidate.get(
                            "name",
                            "",
                        ),
                    )
                    or ""
                )
                .strip()
                .upper()
            )

            if not candidate_name:
                continue

            candidate_party = (
                candidate.get(
                    "pn"
                )
                or candidate.get(
                    "party"
                )
                or candidate.get(
                    "partyName"
                )
                or ""
            )

            party_key = get_party_key(
                candidate_party
            )

            if party_key is None:
                continue

            votes = safe_int(
                candidate.get(
                    "vc",
                    candidate.get(
                        "votes",
                        0,
                    ),
                )
            )

            party_rows[
                party_key
            ][
                candidate_name
            ] = (
                party_rows[
                    party_key
                ].get(
                    candidate_name,
                    0,
                )
                + votes
            )

            party_has_candidate[
                party_key
            ] = True

        # Some Federal payloads put write-ins at the town level.
        write_ins = (
            town_block.get(
                "wc",
                []
            )
            or []
        )

        total_write_ins = sum(
            safe_int(
                item.get(
                    "vc",
                    item.get(
                        "votes",
                        0,
                    ),
                )
            )
            for item in write_ins
        )

        source_total = None

        if "sc" in town_block:

            source_total = safe_int(
                town_block.get(
                    "sc"
                )
            )

        elif "total" in town_block:

            source_total = safe_int(
                town_block.get(
                    "total"
                )
            )

        for party_key in [
            "dem",
            "rep",
        ]:

            if not party_has_candidate[
                party_key
            ]:
                continue

            party_rows[
                party_key
            ][
                "Total Write Ins"
            ] = total_write_ins

            if source_total is not None:

                party_rows[
                    party_key
                ][
                    "Total"
                ] = source_total

            rows_by_party[
                party_key
            ].append(
                party_rows[
                    party_key
                ]
            )

    result = {}

    for party_key, rows in rows_by_party.items():

        if not rows:

            result[
                party_key
            ] = pd.DataFrame()

            continue

        df = pd.DataFrame(
            rows
        )

        text_columns = {
            "Town",
            "Rep District",
            "Sen District",
        }

        for column in df.columns:

            if column in text_columns:

                df[
                    column
                ] = (
                    df[
                        column
                    ]
                    .fillna("")
                    .astype(str)
                )

                continue

            df[
                column
            ] = (
                pd.to_numeric(
                    df[
                        column
                    ],
                    errors="coerce",
                )
                .fillna(0)
                .astype(int)
            )

        aggregation = {}

        for column in df.columns:

            if column == "Town":
                continue

            if column in {
                "Rep District",
                "Sen District",
            }:

                aggregation[
                    column
                ] = combine_unique

            else:

                aggregation[
                    column
                ] = "sum"

        df = (
            df.groupby(
                "Town",
                as_index=False,
            )
            .agg(
                aggregation
            )
        )

        protected = {
            "Town",
            "Rep District",
            "Sen District",
            "Total Write Ins",
            "Others",
            "Total",
        }

        drop_columns = []

        for column in df.columns:

            if column in protected:
                continue

            numeric = pd.to_numeric(
                df[
                    column
                ],
                errors="coerce",
            )

            if (
                numeric.notna().any()
                and numeric
                .fillna(0)
                .sum()
                == 0
            ):

                drop_columns.append(
                    column
                )

        if drop_columns:

            df = df.drop(
                columns=
                    drop_columns
            )

        first_columns = [
            column
            for column in [
                "Town",
                "Rep District",
                "Sen District",
            ]
            if column in df.columns
        ]

        near_end = [
            column
            for column in [
                "Total Write Ins",
                "Others",
            ]
            if column in df.columns
        ]

        candidate_columns = [
            column
            for column in df.columns
            if (
                column not in first_columns
                and column not in near_end
                and column != "Total"
            )
        ]

        ordered = (
            first_columns
            + candidate_columns
            + near_end
        )

        if "Total" in df.columns:

            ordered.append(
                "Total"
            )

        result[
            party_key
        ] = (
            df[
                ordered
            ]
            .sort_values(
                "Town"
            )
            .reset_index(
                drop=True
            )
        )

    return result


def process_federal_json(
    federal_data,
    observed_at,
    baseline_at=None,
):

    """
    Find the live U.S. House office in the Vermont Federal
    payload, build Democratic/Republican town rows, compare
    against prior observations, and write safe CSV files.

    Last Vote Change stays blank on the initial baseline and
    changes only when that town's actual result values change.
    """

    party_blocks = (
        federal_data.get(
            "d",
            []
        )
        or []
    )

    # First try the same party->office structure as Statewide.
    built = {
        "dem": pd.DataFrame(),
        "rep": pd.DataFrame(),
    }

    for party_block in party_blocks:

        party_key = get_party_key(
            party_block.get(
                "pn",
                "",
            )
        )

        offices = (
            party_block.get(
                "o",
                []
            )
            or []
        )

        for office_block in offices:

            office_name = normalize_office_name(
                office_block.get(
                    "on",
                    office_block.get(
                        "officeName",
                        "",
                    ),
                )
            )

            if (
                office_name
                and office_name
                not in FEDERAL_OFFICE_NAMES
                and "REPRESENTATIVE" not in office_name
                and "CONGRESS" not in office_name
            ):
                continue

            # If party is supplied by the outer block and candidate
            # rows do not carry party, copy it into candidates.
            if party_key is not None:

                office_copy = json.loads(
                    json.dumps(
                        office_block
                    )
                )

                for town_block in (
                    office_copy.get(
                        "cs",
                        []
                    )
                    or []
                ):

                    for candidate in (
                        town_block.get(
                            "rc",
                            []
                        )
                        or []
                    ):

                        if not (
                            candidate.get(
                                "pn"
                            )
                            or candidate.get(
                                "party"
                            )
                            or candidate.get(
                                "partyName"
                            )
                        ):

                            candidate[
                                "pn"
                            ] = (
                                "Democratic"
                                if party_key == "dem"
                                else "Republican"
                            )

                office_results = (
                    build_federal_office(
                        office_copy
                    )
                )

            else:

                office_results = (
                    build_federal_office(
                        office_block
                    )
                )

            for key in [
                "dem",
                "rep",
            ]:

                candidate_df = (
                    office_results.get(
                        key
                    )
                )

                if (
                    candidate_df is not None
                    and not candidate_df.empty
                ):

                    if built[
                        key
                    ].empty:

                        built[
                            key
                        ] = candidate_df

                    else:

                        built[
                            key
                        ] = pd.concat(
                            [
                                built[
                                    key
                                ],
                                candidate_df,
                            ],
                            ignore_index=True,
                        )

    # Fallback: Federal payload itself may be one office block.
    if (
        built[
            "dem"
        ].empty
        and built[
            "rep"
        ].empty
    ):

        fallback = build_federal_office(
            federal_data
        )

        built.update(
            fallback
        )

    created = 0

    for party_key, output_file in [
        (
            "dem",
            FEDERAL_DEM_FILE,
        ),
        (
            "rep",
            FEDERAL_REP_FILE,
        ),
    ]:

        df = built.get(
            party_key
        )

        if (
            df is None
            or df.empty
        ):

            print(
                "No live Federal rows found for:",
                party_key,
            )

            continue

        df = add_ru_timestamps(
            df,
            category="federal",
            party=party_key,
            office="REPRESENTATIVE TO CONGRESS",
            observed_at=observed_at,
            baseline_at=baseline_at,
        )

        if "_source_last_updated" in df.columns:

            df = df.drop(
                columns=[
                    "_source_last_updated"
                ]
            )

        df = sort_by_last_updated(
            df
        )

        df = order_timestamp_column(
            df
        )

        atomic_to_csv(
            df,
            output_file,
        )

        print(
            "Created live Federal CSV:",
            output_file,
        )

        print(
            "Rows:",
            len(
                df
            ),
        )

        created += 1

    return created


# =========================================================
# PROCESS STATEWIDE JSON
# =========================================================

def process_statewide_json(
    statewide_data,
    observed_at,
    baseline_at=None,
):

    parties = (
        statewide_data.get(
            "d",
            []
        )
        or []
    )


    created = 0


    print(
        "Statewide party blocks:",
        len(
            parties
        ),
    )


    for party_block in parties:

        party_name = (
            party_block.get(
                "pn",
                ""
            )
        )


        party_key = get_party_key(
            party_name
        )


        if party_key is None:

            print(
                "Skipping party:",
                party_name,
            )

            continue


        offices = (
            party_block.get(
                "o",
                []
            )
            or []
        )


        for office_block in offices:

            office_name = (
                str(
                    office_block.get(
                        "on",
                        ""
                    )
                    or ""
                )
                .strip()
                .upper()
            )


            if not office_name:
                continue


            if (
                office_name
                not in STATEWIDE_FILES
            ):

                print(
                    "Skipping office:",
                    office_name,
                )

                continue


            print(
                "Processing:",
                party_name,
                "-",
                office_name,
            )


            df = build_statewide_office(
                office_block
            )


            if df.empty:

                print(
                    "No rows found:",
                    party_name,
                    "-",
                    office_name,
                )

                continue


            # =============================================
            # ADD PER-RU OBSERVED CHANGE TIMESTAMP
            # =============================================

            df = add_ru_timestamps(
                df,
                category="statewide",
                party=party_key,
                office=office_name,
                observed_at=observed_at,
                baseline_at=baseline_at,
            )


            # =============================================
            # MOST RECENT RU FIRST
            # =============================================

            if "_source_last_updated" in df.columns:

                df = df.drop(
                    columns=[
                        "_source_last_updated"
                    ]
                )

            df = sort_by_last_updated(
                df
            )


            # =============================================
            # SHOW TIMESTAMP FIRST
            # =============================================

            df = order_timestamp_column(
                df
            )


            output_file = (
                STATEWIDE_FILES[
                    office_name
                ][
                    party_key
                ]
            )


            atomic_to_excel(
                df,
                output_file,
            )


            created += 1


            print(
                "Created:",
                output_file,
            )


            print(
                "Rows:",
                len(
                    df
                ),
            )


    return created


# =========================================================
# ENSURE FEDERAL CSV EXISTS
# =========================================================

def ensure_federal_csv(
    csv_path,
    legacy_xlsx_path,
):

    if csv_path.exists():
        return True

    if not legacy_xlsx_path.exists():
        print(
            "Federal source file not found:",
            csv_path.name,
            "or",
            legacy_xlsx_path.name,
        )
        return False

    try:

        df = pd.read_excel(
            legacy_xlsx_path
        )

        atomic_to_csv(
            df,
            csv_path,
        )

        print(
            "Migrated Federal XLSX to CSV:",
            csv_path.name,
        )

        return True

    except Exception as error:

        print(
            "Could not migrate Federal XLSX:",
            legacy_xlsx_path,
            error,
        )

        return False


# =========================================================
# ADD TIMESTAMPS TO EXISTING FEDERAL FILES
# =========================================================

def update_federal_timestamps(
    path,
    legacy_path,
    party,
    observed_at,
):

    if not ensure_federal_csv(
        path,
        legacy_path,
    ):

        return


    try:

        df = pd.read_csv(
            path,
            dtype=str,
        )

    except Exception as error:

        print(
            "Could not timestamp Federal file:",
            path,
            error,
        )

        return


    if (
        df.empty
        or "Town" not in df.columns
    ):

        return


    # Remove a previous timestamp before calculating
    # signatures so the timestamp itself never causes
    # a false result change.

    if "Last Updated" in df.columns:

        df = df.drop(
            columns=[
                "Last Updated"
            ]
        )


    df = add_ru_timestamps(
        df,
        category="federal",
        party=party,
        office="REPRESENTATIVE TO CONGRESS",
        observed_at=observed_at,
    )


    df = sort_by_last_updated(
        df
    )


    df = order_timestamp_column(
        df
    )


    atomic_to_csv(
        df,
        path,
    )


    print(
        "Updated Federal RU timestamps:",
        path.name,
    )


# =========================================================
# CHECK RESULTS
# =========================================================

async def check_results():

    checked_at = now_string()


    print()
    print(
        "=" * 70
    )


    print(
        "Checking Vermont results:",
        checked_at,
    )


    async with httpx.AsyncClient(
        timeout=30.0,
        follow_redirects=True,
        headers={
            "User-Agent":
                "Vermont Election Results Monitor"
        },
    ) as client:


        # =================================================
        # MANIFEST
        # =================================================

        manifest = await fetch_json(
            client,
            MANIFEST_URL,
        )


        # =================================================
        # REPORTING UNITS
        # =================================================

        reporting = manifest.get(
            "townsReporting",
            "0/247",
        )


        try:

            reported_text, total_text = (
                reporting.split(
                    "/"
                )
            )


            reported = int(
                reported_text.strip()
            )


            total_units = int(
                total_text.strip()
            )


        except Exception:

            reported = 0
            total_units = 247


        vermont_last_updated = (
            source_time_to_eastern(
                manifest.get(
                    "lastUpdatedDate"
                )
            )
        )


        election_status = (
            get_election_status(
                manifest
            )
        )


        print(
            "Reporting:",
            f"{reported}/{total_units}",
        )


        print(
            "Vermont last updated:",
            vermont_last_updated,
        )


        print(
            "Election status:",
            election_status,
        )


        # =================================================
        # CURRENT DATA PATHS
        # =================================================

        federal_path = (
            manifest
            .get(
                "federal",
                {},
            )
            .get(
                "path"
            )
        )


        statewide_path = (
            manifest
            .get(
                "stateWide",
                {},
            )
            .get(
                "path"
            )
        )


        federal_url = build_static_url(
            federal_path
        )


        statewide_url = build_static_url(
            statewide_path
        )


        print(
            "Federal JSON:",
            federal_url,
        )


        print(
            "Statewide JSON:",
            statewide_url,
        )


        # =================================================
        # STATEWIDE
        # =================================================

        statewide_created = 0


        if statewide_url:

            statewide_data = await fetch_json(
                client,
                statewide_url,
            )


            atomic_write_text(
                STATEWIDE_RAW_JSON,
                json.dumps(
                    statewide_data,
                    indent=2,
                ),
            )


            statewide_created = (
                process_statewide_json(
                    statewide_data,
                    observed_at=checked_at,
                    baseline_at=vermont_last_updated,
                )
            )


        else:

            print(
                "No Statewide path found."
            )


        # =================================================
        # FEDERAL
        # =================================================

        federal_created = 0

        if federal_url:

            federal_data = await fetch_json(
                client,
                federal_url,
            )

            federal_created = (
                process_federal_json(
                    federal_data,
                    observed_at=checked_at,
                    baseline_at=vermont_last_updated,
                )
            )

        else:

            print(
                "No Federal path found."
            )


        # =================================================
        # STATUS.JSON
        # =================================================

        save_status(

            last_checked=
                checked_at,

            last_updated=
                vermont_last_updated,

            election_status=
                election_status,


            # Federal

            federal_dem_reporting=
                reported,

            federal_rep_reporting=
                reported,

            federal_overall_reporting=
                reported,

            federal_total_units=
                total_units,

            federal_results_path=
                federal_path,

            federal_files_created=
                federal_created,


            # Statewide

            statewide_dem_reporting=
                reported,

            statewide_rep_reporting=
                reported,

            statewide_overall_reporting=
                reported,

            statewide_total_units=
                total_units,

            statewide_results_path=
                statewide_path,

            statewide_files_created=
                statewide_created,


            error=None,
        )


        print()
        print(
            "Statewide files created:",
            statewide_created,
        )


# =========================================================
# MONITOR LOOP
# =========================================================

async def monitor():

    print(
        "Starting Vermont election monitor."
    )


    print(
        "Election GUID:",
        ELECTION_GUID,
    )


    print(
        "Manifest:",
        MANIFEST_URL,
    )


    print(
        "Checking every",
        CHECK_INTERVAL,
        "seconds.",
    )


    while True:

        try:

            await check_results()


        except KeyboardInterrupt:

            raise


        except Exception as error:

            message = str(
                error
            )


            print()
            print(
                "MONITOR ERROR:",
                message,
            )


            save_status(
                last_checked=
                    now_string(),

                error=
                    message,
            )


        print()
        print(
            "Sleeping",
            CHECK_INTERVAL,
            "seconds..."
        )


        await asyncio.sleep(
            CHECK_INTERVAL
        )


# =========================================================
# START
# =========================================================

if __name__ == "__main__":

    try:

        asyncio.run(
            monitor()
        )


    except KeyboardInterrupt:

        print()
        print(
            "Monitor stopped."
        )