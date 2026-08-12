# To run:
# uv run python monitor.py

from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import asyncio
import json


import httpx
import pandas as pd


# =========================================================
# SETTINGS
# =========================================================

APP_DIR = Path(__file__).resolve().parent

RESULTS_DIR = APP_DIR / "vermont_results"
FEDERAL_DIR = RESULTS_DIR / "federal"

STATEWIDE_DIR = RESULTS_DIR / "statewide"
STATEWIDE_DEM_DIR = STATEWIDE_DIR / "democrat"
STATEWIDE_REP_DIR = STATEWIDE_DIR / "republican"

STATUS_FILE = RESULTS_DIR / "status.json"

STATEWIDE_RAW_JSON = (
    STATEWIDE_DIR / "statewide_live.json"
)

CHECK_INTERVAL = 30


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
        ZoneInfo("America/New_York")
    ).strftime(
        "%m/%d/%Y %I:%M:%S %p"
    )


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

    STATUS_FILE.write_text(
        json.dumps(
            status,
            indent=2,
        )
    )


# =========================================================
# HELPERS
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


        row = {
            "Town":
                town,

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
        #
        # Vermont uses wc for write-in candidates.
        #
        # Every wc vote gets added to Total Write Ins.
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

            write_in_votes = (
                safe_int(
                    write_in.get(
                        "vc"
                    )
                )
            )

            total_write_ins += (
                write_in_votes
            )


        row[
            "Total Write Ins"
        ] = total_write_ins


        # =================================================
        # OTHERS
        #
        # Keep if Vermont supplies an explicit field.
        # =================================================

        others = None


        for key in [
            "others",
            "Others",
            "oc",
        ]:

            if key in town_block:

                others = (
                    safe_int(
                        town_block.get(
                            key
                        )
                    )
                )

                break


        if others is not None:

            row[
                "Others"
            ] = others


        # =================================================
        # TOTAL
        #
        # sc = Vermont's official source total.
        #
        # Do not manufacture this ourselves.
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
    # FILL NUMERIC COLUMNS
    # =====================================================

    text_columns = {
        "Town",
        "Rep District",
        "Sen District",
    }


    for column in df.columns:

        if column in text_columns:

            df[column] = (
                df[column]
                .fillna("")
                .astype(str)
            )

            continue


        df[column] = (
            pd.to_numeric(
                df[column],
                errors="coerce",
            )
            .fillna(0)
            .astype(int)
        )


    # =====================================================
    # COMBINE DUPLICATE TOWNS
    #
    # Example:
    #
    # BENNINGTON | BENNINGTON-2 | ... | 820
    # BENNINGTON | BENNINGTON-5 | ... | 625
    #
    # becomes:
    #
    # BENNINGTON | BENNINGTON-2, BENNINGTON-5 | ... | 1445
    # =====================================================

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


    # =====================================================
    # REMOVE ZERO-ONLY CANDIDATES
    #
    # Keep Total Write Ins even if zero.
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
            df[column],
            errors="coerce",
        )


        if (
            numeric.notna().any()
            and
            numeric.fillna(0).sum() == 0
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
    #
    # Total always last.
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

        for column
        in df.columns

        if (
            column
            not in first_columns

            and

            column
            not in near_end

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
# PROCESS STATEWIDE JSON
# =========================================================

def process_statewide_json(
    statewide_data,
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


        party_key = (
            get_party_key(
                party_name
            )
        )


        if party_key is None:

            print(
                "Skipping party:",
                party_name,
            )

            continue


        print()

        print(
            "Processing party:",
            party_name,
        )


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


            df = (
                build_statewide_office(
                    office_block
                )
            )


            if df.empty:

                print(
                    "No rows found:",
                    party_name,
                    "-",
                    office_name,
                )

                continue


            output_file = (
                STATEWIDE_FILES[
                    office_name
                ][
                    party_key
                ]
            )


            df.to_excel(
                output_file,
                index=False,
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


            print(
                "Columns:",
                list(
                    df.columns
                ),
            )


    return created


# =========================================================
# CHECK RESULTS
# =========================================================

async def check_results():

    checked_at = (
        now_string()
    )


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

        manifest = (
            await fetch_json(
                client,
                MANIFEST_URL,
            )
        )


        # =================================================
        # REPORTING UNITS
        # =================================================

        reporting = (
            manifest.get(
                "townsReporting",
                "0/247",
            )
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


        last_updated = (
            manifest.get(
                "lastUpdatedDate"
            )
        )


        print(
            "Reporting:",
            f"{reported}/{total_units}",
        )


        print(
            "Last updated:",
            last_updated,
        )


        # =================================================
        # FEDERAL PATH
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


        # =================================================
        # STATEWIDE PATH
        # =================================================

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


        federal_url = (
            build_static_url(
                federal_path
            )
        )


        statewide_url = (
            build_static_url(
                statewide_path
            )
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
        # DOWNLOAD + PROCESS STATEWIDE
        # =================================================

        statewide_created = 0


        if statewide_url:

            statewide_data = (
                await fetch_json(
                    client,
                    statewide_url,
                )
            )


            # Save exact source JSON locally.
            STATEWIDE_RAW_JSON.write_text(
                json.dumps(
                    statewide_data,
                    indent=2,
                )
            )


            statewide_created = (
                process_statewide_json(
                    statewide_data
                )
            )


        else:

            print(
                "No Statewide path found."
            )


        # =================================================
        # STATUS.JSON
        # =================================================

        save_status(

            last_checked=
                checked_at,

            last_updated=
                last_updated,


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