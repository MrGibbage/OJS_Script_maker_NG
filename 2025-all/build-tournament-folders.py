"""Utility to prepare per-tournament folders and populate OJS spreadsheets.

This script reads a season manifest (`season.json`) and a master
`TournamentList` (or `DivTournamentList`) workbook to create a folder for
each tournament, copy template files, and populate the OJS (Online Judge
System) spreadsheet tables with team/award/meta information.

Helpers use `openpyxl` to manipulate tables inside each OJS file and
`pandas` for convenient table-level I/O. Several functions try to preserve or
replicate workbook features such as data validation, conditional formatting,
formulas and cell styles when expanding table rows.

The script uses best-effort error handling so a single malformed cell or
formatting object does not abort the whole run.
"""

# To create a windows executable file,
# first install pyinstaller, with
# pip install pyinstaller
# then run
# .venv\Scripts\pyinstaller.exe -F 2025\build-tournament-folders.py
# Then copy the build-tournament-folders.exe file from dist to 2025
#
import os, sys, re, time
import shutil
import warnings
import configparser

# pip install openpyxl
from openpyxl import load_workbook
from openpyxl.utils.cell import (
    column_index_from_string,
    get_column_letter,
    coordinate_from_string,
)
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import Rule, CellIsRule, FormulaRule
from copy import copy as _copy
from openpyxl.worksheet.datavalidation import DataValidation

# pip install print-color
from colorama import Fore, Back, Style, init

# pip install pywin32
import win32com.client

# pip install pandas
import pandas as pd

# pip install xlwings
import xlwings

import math
import numpy as np

SHEET_PASSWORD: str = "skip"

import json
from typing import Union, Dict, Any, Sequence


def print_error(errormsg: str, e: Exception):
    print(Fore.RED)
    print(f"{errormsg}\n{e}")
    input("Press enter to quit...")
    print(Fore.RESET)
    sys.exit(1)


def _copy_data_validations_for_range(
    ws: Worksheet,
    start_col_idx: int,
    end_col_idx: int,
    first_data_row: int,
    new_end_row: int,
):
    """
    Duplicate any data validations that apply to the template data row (first_data_row)
    and add the same validation for the full new range first_data_row..new_end_row across
    only the original validation's columns intersected with start_col_idx..end_col_idx.
    """
    try:
        existing = list(ws.data_validations.dataValidation)
    except Exception:
        existing = []

    for dv in existing:
        try:
            # iterate each range the validation currently applies to
            for rng in dv.ranges:
                try:
                    cr = CellRange(str(rng))
                except Exception:
                    continue

                # Only consider ranges that include the template data row
                if not (cr.min_row <= first_data_row <= cr.max_row):
                    continue

                # Determine intersection of columns between the original dv range and our target table columns
                orig_min_col = cr.min_col
                orig_max_col = cr.max_col
                new_min_col = max(orig_min_col, start_col_idx)
                new_max_col = min(orig_max_col, end_col_idx)

                if new_min_col > new_max_col:
                    # no column overlap with this table
                    continue

                # build the new target range for the intersection columns
                new_range = f"{get_column_letter(new_min_col)}{first_data_row}:{get_column_letter(new_max_col)}{new_end_row}"

                # create a new DataValidation copying core properties
                newdv = DataValidation(
                    type=dv.type,
                    formula1=getattr(dv, "formula1", None),
                    formula2=getattr(dv, "formula2", None),
                    allow_blank=getattr(dv, "allow_blank", None),
                    operator=getattr(dv, "operator", None),
                    showDropDown=getattr(dv, "showDropDown", None),
                    error=getattr(dv, "error", None),
                    errorTitle=getattr(dv, "errorTitle", None),
                    prompt=getattr(dv, "prompt", None),
                    promptTitle=getattr(dv, "promptTitle", None),
                )

                # add the new range and attach to worksheet
                try:
                    newdv.add(new_range)
                    ws.add_data_validation(newdv)
                except Exception:
                    # best-effort: skip if we cannot add this cloned validation
                    continue
        except Exception:
            # best-effort: don't abort on any unexpected validation object shape
            continue


def _extend_conditional_formatting_for_range(
    ws: Worksheet,
    start_col_letter: str,
    end_col_letter: str,
    first_data_row: int,
    new_end_row: int,
):
    """
    Duplicate conditional-formatting rules that cover the template data row
    and add equivalent rules for the new range.
    """
    try:
        cf_rules = getattr(ws.conditional_formatting, "_cf_rules", {})
    except Exception:
        cf_rules = {}

    new_range = f"{start_col_letter}{first_data_row}:{end_col_letter}{new_end_row}"

    # cf_rules keys may be space-separated ranges; iterate safely
    for key, rules in list(cf_rules.items()):
        try:
            # key may be e.g. 'A2:A10' or 'A2:A10 B2:B10' -> split and examine each
            for sub in str(key).split():
                try:
                    cr = CellRange(sub)
                except Exception:
                    continue
                # if the rule applied to the template data row, clone rules for new_range
                if cr.min_row <= first_data_row <= cr.max_row:
                    for rule in rules:
                        try:
                            ws.conditional_formatting.add(new_range, rule)
                        except Exception:
                            # some rule objects may not be directly re-addable; skip if so
                            continue
        except Exception:
            continue


def _to_int(val: Any, default: int = 0) -> int:
    """
    Safely coerce val to an int.
    Accepts:
      - pandas.Series (uses first element)
      - numpy arrays (uses first element)
      - lists/tuples (uses first element)
      - scalars (returned as int if possible)
    Returns `default` on missing/NaN/uncoercible values.
    """
    try:
        # pandas Series -> first element
        if isinstance(val, pd.Series):
            if val.empty:
                return default
            v = val.iat[0]
        # numpy array -> first element
        elif isinstance(val, np.ndarray):
            if val.size == 0:
                return default
            v = val.flat[0]
        # sequence -> first element
        elif isinstance(val, (list, tuple)):
            if len(val) == 0:
                return default
            v = val[0]
        else:
            v = val

        if pd.isna(v):
            return default
        return int(v)
    except Exception:
        return default


def add_table_dataframe(
    wb: Workbook,
    sheet_name: str,
    table_name: str,
    data: pd.DataFrame,
    require_all_columns: bool = True,
    keep_vba: bool = True,
) -> int:
    """
    Append a pandas.DataFrame to an existing Excel table.
    - Validates that the DataFrame columns match the table headers (exact order by default).
    - Fills any completely-blank data rows inside the table before appending new rows.
    - Updates table.ref if rows are appended.
    Returns the number of rows written.
    """
    if data is None or data.empty:
        print_error("Trying to add_table_dataframe, but data is none")
        return 0
    else:
        print("Here's the dataframe:")
        print(data)

    if sheet_name not in wb.sheetnames:
        print_error(f"{sheet_name} worksheet not found")
        raise KeyError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]

    if table_name not in ws.tables:
        print_error(f"{table_name} table not found")
        raise KeyError(f"Table {table_name!r} not found on sheet {sheet_name!r}")
    table = ws.tables[table_name]
    table_range = table.ref  # e.g. "C2:F20"

    # header row and data rows within the table ref
    table_head = ws[table_range][0]
    table_data_rows = ws[table_range][1:]
    headers = [
        c.value.strip() if isinstance(c.value, str) else c.value for c in table_head
    ]

    # Normalize DataFrame column names (strip strings)
    df = data.copy()
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

    # Validate columns match
    if require_all_columns:
        if list(df.columns) != headers:
            raise ValueError(
                f"DataFrame columns do not match table headers for {table_name!r} on sheet {sheet_name!r}.\n"
                f"Table headers: {headers}\nDataFrame columns: {list(df.columns)}"
            )

    start_cell, end_cell = table_range.split(":")
    _, end_row = coordinate_from_string(end_cell)
    start_col_letter, start_row = coordinate_from_string(start_cell)
    start_col_idx = column_index_from_string(start_col_letter)

    rows_written = 0
    df_iter_index = 0
    total_rows = len(df)

    # Fill existing blank data rows first
    for row_tuple in table_data_rows:
        if df_iter_index >= total_rows:
            print("about to break")
            break
        # consider a row blank if all cells are None or blank strings
        if all(
            (cell.value is None)
            or (isinstance(cell.value, str) and cell.value.strip() == "")
            for cell in row_tuple
        ):
            target_row_idx = row_tuple[0].row
            row_values = df.iloc[df_iter_index]
            for j, col_name in enumerate(headers):
                val = row_values[col_name] if col_name in df.columns else None
                ws.cell(row=target_row_idx, column=start_col_idx + j).value = val
            df_iter_index += 1
            rows_written += 1

    # Append remaining rows after end_row
    current_row = int(end_row)
    appended = False
    while df_iter_index < total_rows:
        current_row += 1
        row_values = df.iloc[df_iter_index]
        for j, col_name in enumerate(headers):
            val = row_values[col_name] if col_name in df.columns else None
            ws.cell(row=current_row, column=start_col_idx + j).value = val
        df_iter_index += 1
        rows_written += 1
        appended = True

    # If we appended rows, extend the table ref to include the new end row
    if appended:
        end_col_letter = coordinate_from_string(end_cell)[0]
        table.ref = f"{start_cell}:{end_col_letter}{current_row}"

    return rows_written


def read_table_as_df(
    xlsx_path: str,
    sheet_name: str,
    table_name: str,
    require_table: bool = True,
    convert_integer_floats: bool = True,
) -> pd.DataFrame:
    """
    Read an Excel table (ListObject) by name into a pandas DataFrame.

    Same behaviour as before, with an optional post-processing step:
    - If convert_integer_floats is True (default), any float column where all
      non-missing values are whole numbers (e.g. 1.0, 2.0) will be converted
      to the pandas nullable integer dtype "Int64" so integers are preserved
      while keeping NA support.
    """
    try:
        wb = load_workbook(xlsx_path, data_only=True)
    except Exception as e:
        raise RuntimeError(f"Failed to open workbook '{xlsx_path}': {e}") from e

    if sheet_name not in wb.sheetnames:
        if require_table:
            raise KeyError(f"Sheet not found: {sheet_name} in {xlsx_path}")
        return pd.DataFrame()

    ws = wb[sheet_name]

    if table_name not in ws.tables:
        if require_table:
            raise KeyError(
                f"Table {table_name!r} not found on sheet {sheet_name!r} in {xlsx_path}"
            )
        return pd.DataFrame()

    table = ws.tables[table_name]
    ref = table.ref  # expected e.g. "C2:F20"
    if not isinstance(ref, str) or ":" not in ref:
        raise ValueError(
            f"Unexpected table.ref for {table_name!r} on {sheet_name!r}: {ref!r}"
        )

    try:
        start, end = ref.split(":")
        start_col, start_row = coordinate_from_string(start)
        end_col, end_row = coordinate_from_string(end)
    except Exception as e:
        raise ValueError(
            f"Could not parse table.ref '{ref}' for {table_name!r}: {e}"
        ) from e

    header_row_idx = int(start_row) - 1  # pandas header is 0-based
    usecols = f"{start_col}:{end_col}"
    nrows = int(end_row) - int(start_row)  # number of data rows after header
    if nrows <= 0:
        # empty table body -> return empty DataFrame with headers read if possible
        try:
            df = pd.read_excel(
                xlsx_path,
                sheet_name=sheet_name,
                header=header_row_idx,
                usecols=usecols,
                nrows=0,
                engine="openpyxl",
            )
            # still perform cleanup on the empty df
            df.columns = df.columns.str.strip()
            return df
        except Exception:
            return pd.DataFrame()

    try:
        df = pd.read_excel(
            xlsx_path,
            sheet_name=sheet_name,
            header=header_row_idx,
            usecols=usecols,
            nrows=nrows,
            engine="openpyxl",
        )
    except Exception as e:
        raise RuntimeError(
            f"pandas.read_excel failed for table {table_name!r} on sheet {sheet_name!r}: {e}"
        ) from e

    # basic cleanup: strip column names and trim string values safely (vectorized where possible)
    df.columns = df.columns.str.strip()

    def _trim_series(s: pd.Series) -> pd.Series:
        if pd.api.types.is_string_dtype(s):
            return s.str.strip()
        if s.dtype == object:
            return s.map(lambda v: v.strip() if isinstance(v, str) else v)
        return s

    df = df.apply(_trim_series)

    # NEW: convert float columns that are integer-like into pandas nullable Int64 dtype
    if convert_integer_floats:
        float_cols = df.select_dtypes(include=["float"]).columns
        for col in float_cols:
            ser = df[col]
            non_na = ser.dropna()
            if non_na.empty:
                # nothing to convert (all NA) -> skip conversion to avoid choosing int dtype when ambiguous
                continue
            # vectorized check: are all non-NA values whole numbers?
            try:
                # use modulo 1 check; for numerical stability cast to float
                if ((non_na % 1) == 0).all():
                    # convert to pandas nullable integer dtype to preserve NA
                    df[col] = df[col].astype("Int64")
            except Exception:
                # best-effort: if any operation fails, skip conversion for this column
                continue

    return df


def read_table_as_dict(
    xlsx_path: str,
    sheet_name: str,
    table_name: str,
    key_col: str | None = None,
    value_col: str | None = None,
    require_unique_keys: bool = True,
) -> dict:
    """
    Read a two-column Excel table (ListObject) and return a dict mapping key->value.

    Enforces that the table contains exactly two columns. By default raises on duplicate keys;
    set require_unique_keys=False to keep last-seen value for duplicate keys.

    Parameters:
    - xlsx_path: path to workbook
    - sheet_name: sheet containing the table
    - table_name: Excel table (ListObject) name
    - key_col: optional column name to use as key (must exist in table); defaults to first column
    - value_col: optional column name to use as value (must exist in table); defaults to second column
    - require_unique_keys: if True raise ValueError when duplicate keys are encountered

    Returns:
    - dict mapping key -> value
    """
    df = read_table_as_df(xlsx_path, sheet_name, table_name, require_table=True)

    if df.shape[1] != 2:
        raise ValueError(
            f"Table {table_name!r} on sheet {sheet_name!r} must have exactly 2 columns (found {df.shape[1]})"
        )

    # Determine which columns to use
    col_names = list(df.columns)
    key_col_name = key_col if key_col is not None else col_names[0]
    value_col_name = value_col if value_col is not None else col_names[1]

    if key_col_name not in df.columns or value_col_name not in df.columns:
        raise KeyError(
            f"Specified key/value columns not found in table columns: {df.columns.tolist()}"
        )

    mapping: dict = {}
    for idx, row in df.iterrows():
        raw_key = row[key_col_name]
        raw_val = row[value_col_name]

        # skip rows with empty keys
        if pd.isna(raw_key):
            continue

        # normalize strings
        key = raw_key.strip() if isinstance(raw_key, str) else raw_key
        val = raw_val.strip() if isinstance(raw_val, str) else raw_val

        if require_unique_keys and key in mapping:
            raise ValueError(
                f"Duplicate key found in table {table_name!r} on sheet {sheet_name!r}: {key!r}"
            )
        mapping[key] = val

    return mapping


def _remove_note_keys(obj: Any) -> Any:
    """Recursively strip out any dict keys beginning with 'note' (case-ins).

    This is used to allow authors of JSON config files to include free-form
    'note' keys for human-readable comments; they will be removed before the
    data structure is consumed by the rest of the script.
    """
    if isinstance(obj, dict):
        out = {}
        for k, v in obj.items():
            if k.lower().startswith("note"):
                continue
            out[k] = _remove_note_keys(v)
        return out
    if isinstance(obj, list):
        return [_remove_note_keys(i) for i in obj]
    return obj


def load_json_without_notes(path: str) -> dict:
    """
    Load JSON from path and return a copy with any keys starting with "note"
    (case-insensitive) removed at all nesting levels.
    """
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except FileNotFoundError:
        raise
    except json.JSONDecodeError:
        raise
    return _remove_note_keys(data)


# creates the subfolders for each tournament
def create_folder(newpath):
    """Create `newpath` if it does not exist.

    This is a thin helper that prints success and exits via `print_error`
    if directory creation fails for any reason (permission, path invalid,
    etc.). Keeping a single helper centralizes the messaging style.
    """
    if os.path.exists(newpath):
        return
    try:
        os.makedirs(newpath)
        print(Fore.RESET)
        print(f"Created folder: {newpath}")
    except Exception as e:
        print_error(f"Could not create directory: {newpath}", e)


# copies the files into the tournament folders
def copy_files(item: pd.Series):
    """Copy extra files and the OJS template into the tournament folder.

    The `item` parameter is a row from the tournaments DataFrame and is
    expected to contain a `Short Name` (folder name) and an `OJS_FileName`.
    Global variables used:
      - `extrafilelist`: list of filenames to copy into each tournament folder
      - `template_file`: path to the OJS template workbook to copy
      - `dir_path`: base directory containing `tournaments/`

    Any copy error is treated as fatal using `print_error` so the operator can
    inspect and fix missing files before re-running.
    """
    print(item)
    for filename in extrafilelist:
        try:
            # destination folder for this tournament
            newpath = dir_path + "\\tournaments\\" + item["Short Name"] + "\\"
            shutil.copy(dir_path + "\\" + filename, newpath)
        except Exception as e:
            print_error(f"Could not copy file: {filename} to {newpath}", e)

    # Copy the OJS template into place using the filename specified in the
    # tournament list (this may vary per-division or per-tournament).
    try:
        new_ojs_file = (
            dir_path
            + "\\tournaments\\"
            + item["Short Name"]
            + "\\"
            + item["OJS_FileName"]
        )
        shutil.copy(
            template_file,
            new_ojs_file,
        )
    except Exception as e:
        print_error(f"Could not copy OJS file: {template_file} to \\{new_ojs_file}", e)


# edits the OJS spreadsheets with the correct tournament information
def set_up_tapi_worksheet(tournament: pd.Series, book: Workbook):
    """Populate the 'Team and Program Information' table in the open workbook.

    - `tournament` is a pandas Series representing the tournament row.
    - `book` is an open openpyxl `Workbook` object for the tournament's OJS file.

    This function gathers the assigned teams from the global `dfAssignments`
    table, normalizes the columns to the expected TAPI format and then calls
    `add_table_dataframe` to overwrite or extend the `OfficialTeamList` table
    in the workbook.
    """
    # open the OJS workbook
    # print(type(tournament[d + "_OJS"]), tournament[d + "_OJS"])
    d = ""
    print(Fore.RESET)
    if using_divisions:
        d = tournament["Div"]
        if isinstance(tournament["OJS_FileName"], float):
            print_error("Can't get the OJS File name from the tournament spreadsheet")
        print(
            f"Setting up Tournament Team and Program Information for {tournament["Short Name"]} {d}"
        )
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments["Short Name"] == tournament["Short Name"])
            & (dfAssignments["Div"] == d)
        ]
        print(
            f'There are {len(assignees)} teams in this {tournament["Short Name"]} {d} tournament'
        )
    else:
        if isinstance(tournament["OJS_FileName"], float):
            print_error("Can't get the OJS File name from the tournament spreadsheet")
        print(
            f"Setting up Tournament Team and Program Information for {tournament["Short Name"]}"
        )
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments["Short Name"] == tournament["Short Name"])
        ]
        print(
            f'There are {len(assignees)} teams in this {tournament["Short Name"]} tournament'
        )

    # format the assignees df so it matches the TAPI list
    keep = ["Team #", "Team Name", "Coach Name"]
    keep_safe = [c for c in keep if c in assignees.columns]
    assignees = assignees[keep_safe]
    assignees["Pod Number"] = 0
    add_table_dataframe(
        book,
        "Team and Program Information",
        "OfficialTeamList",
        assignees.sort_values(by="Team #", ascending=True),
    )


def set_up_award_worksheet(tournament: pd.Series, book: Workbook):
    """Prepare award tables used by OJS dropdowns and closing scripts.

    Reads award counts from the tournament row, looks up award labels from
    `dfAwardDef` and writes two tables on the `AwardListDropdowns` sheet:
      - `RobotGameAwards`: robot-game specific labels (kept separate)
      - `AwardListDropdowns`: other judged awards used for dropdowns and scripts

    The function tolerates missing label cells and treats non-numeric counts
    as zero using `_to_int`.
    """
    # This will read the awards allocations for each tournament (at the tournament
    # level--no divisions) and create the AwardList table which goes on the
    # normally hidden AwardList worksheet. This list is used to create the
    # dropdowns in the award column on the OJS
    thisDiv = tournament["Div"] if using_divisions else ""
    print(Fore.RESET)
    print(
        f"Setting up Tournament Division Awards for {tournament['Short Name']} {thisDiv}"
    )
    print(tournament)

    # Robot Game
    rg_raw = (
        tournament.get("P_AWD_RG")
        if hasattr(tournament, "get")
        else tournament["P_AWD_RG"]
    )
    try:
        rg_awards = int(rg_raw) if not pd.isna(rg_raw) else 0
    except Exception as e:
        print_error(f"Could not get the number of robot game awards: {e}")
        rg_awards = 0
    print(f"There are {rg_awards} robot game awards")
    rg_awards_df = pd.DataFrame(columns=["Robot Game Awards"])
    for rg_award_num in range(1, rg_awards + 1):
        thisLabel = "Label" + str(rg_award_num)
        # select the matching cell; loc with a boolean mask returns a Series,
        # so extract the first element safely (or None if missing)
        sel = dfAwardDef.loc[dfAwardDef["ColumnName"] == "P_AWD_RG", thisLabel]
        try:
            thisValue = sel.iat[0]  # first scalar value from the Series
        except Exception:
            # fallback: if sel is already a scalar or empty, handle gracefully
            thisValue = sel if not (hasattr(sel, "__len__") and len(sel) == 0) else None
        rg_awards_df.loc[len(rg_awards_df)] = [thisValue]

    # The sheet name is AwardListDropdowns, but there are two tables on that sheet: one
    # is the actual awards dropdown list, but the other is for robot game. We don't
    # put the RG awards in the dropdown list because they are selected automatically
    # based on score. The list is used for the closing ceremony script, so it is
    # still needed.
    add_table_dataframe(book, "AwardListDropdowns", "RobotGameAwards", rg_awards_df)

    # Other judged awards
    j_cols = tournament.filter(regex=f"^J_")
    # print(str(config).replace("'", '"'))
    j_awards_df = pd.DataFrame(columns=["Award"])
    for this_col_name, series in j_cols.items():
        count = _to_int(series)
        for award_num in range(1, count + 1):
            # print(find_item_by_col_name(d=config, col_name=this_col_name))
            label_col = "Label" + str(award_num)
            sel = dfAwardDef.loc[dfAwardDef["ColumnName"] == this_col_name, label_col]
            try:
                thisValue = sel.iat[0]
            except Exception:
                thisValue = (
                    sel if not (hasattr(sel, "__len__") and len(sel) == 0) else None
                )
            pos = len(j_awards_df)
            j_awards_df.loc[pos] = [thisValue]
    add_table_dataframe(book, "AwardListDropdowns", "AwardListDropdowns", j_awards_df)


def set_up_meta_worksheet(tournament: pd.Series, book: Workbook):
    print(Fore.RESET)
    d = tournament["Div"] if using_divisions else ""
    print(f"Setting up meta worksheet for {tournament["Short Name"]} {d}")
    print(tournament)
    scriptfile = (
        dir_path
        + "\\tournaments\\"
        + tournament["Short Name"]
        + "\\closing_ceremony.html"
    )
    dfMeta = pd.DataFrame(columns=["Key", "Value"])
    dfMeta.loc[len(dfMeta)] = {"Key": "Tournament Year", "Value": config["season_yr"]}
    dfMeta.loc[len(dfMeta)] = {
        "Key": "FLL Season Title",
        "Value": config["season_name"],
    }
    dfMeta.loc[len(dfMeta)] = {
        "Key": "Tournament Short Name",
        "Value": tournament["Short Name"],
    }
    dfMeta.loc[len(dfMeta)] = {
        "Key": "Tournament Long Name",
        "Value": tournament["Long Name"],
    }
    dfMeta.loc[len(dfMeta)] = {"Key": "Completed Script File", "Value": scriptfile}
    dfMeta.loc[len(dfMeta)] = {"Key": "Using Divisions", "Value": using_divisions}
    if using_divisions:
        dfMeta.loc[len(dfMeta)] = {"Key": "Division", "Value": d}
    dfMeta.loc[len(dfMeta)] = {"Key": "Advancing", "Value": tournament["ADV"]}

    # Add Meta into the open workbook supplied by the caller (book)
    if tournament["OJS_FileName"] is not None:
        add_table_dataframe(book, "Meta", "Meta", dfMeta)


def copy_team_numbers(
    source_sheet: Worksheet,
    target_sheet: Worksheet,
    target_start_row: int,
    source_start_row: int = 3,
    debug: bool = False,
) -> int:
    """
    Copy team numbers from column A of source_sheet to column A of target_sheet.
    - Copies from source_start_row .. last non-empty row (inclusive).
    - Writes starting at target_start_row (increments for each copied row).
    - Returns number of rows copied.

    Set debug=True to print diagnostic info about what's read and written.
    """
    col = 1  # column A
    max_row = source_sheet.max_row

    if debug:
        print(Fore.RESET)
        print(
            f"copy_team_numbers: source_sheet={source_sheet.title}, target_sheet={target_sheet.title}"
        )
        print(
            f"max_row={max_row}, source_start_row={source_start_row}, target_start_row={target_start_row}"
        )

    # Find the last non-empty row in column A
    last_row = None
    for r in range(source_start_row, max_row + 1):
        v = source_sheet.cell(row=r, column=col).value
        if debug:
            print(f"read source row {r}: {repr(v)}")
        if v is not None and v != "":
            last_row = r

    if last_row is None:
        if debug:
            print("copy_team_numbers: no non-empty rows found in source column A")
        return 0

    team_count = last_row - source_start_row + 1
    if debug:
        print(f"Detected last_row={last_row}, team_count={team_count}")

    # Copy inclusive from source_start_row to last_row
    dest_row = target_start_row
    copied = 0
    for r in range(source_start_row, last_row + 1):
        cell_value = source_sheet.cell(row=r, column=col).value
        if debug:
            print(
                f"writing to target row {dest_row}: {repr(cell_value)} (from source row {r})"
            )
        # normalize numpy / pandas numeric types to native Python types before writing
        if pd.isna(cell_value):
            write_val = None
        elif isinstance(cell_value, (np.integer,)):
            write_val = int(cell_value)
        elif isinstance(cell_value, (np.floating,)):
            fv = float(cell_value)
            write_val = int(fv) if fv.is_integer() else fv
        else:
            write_val = cell_value

        target_sheet.cell(row=dest_row, column=col).value = write_val
        dest_row += 1
        copied += 1

    if debug:
        print(f"copy_team_numbers: finished, copied={copied}")

    return copied


def protect_worksheets(tournament: pd.Series, book: Workbook):
    """Apply protection settings to every worksheet in `book`.

    Uses the module-level `SHEET_PASSWORD`. The protection is conservative and
    disables formatting and structural changes while allowing only selection of
    locked cells.
    """
    for ws in book.worksheets:
        ws.protection.selectLockedCells = True
        ws.protection.selectUnlockedCells = False
        ws.protection.formatCells = False
        ws.protection.formatColumns = False
        ws.protection.formatRows = False
        ws.protection.autoFilter = False
        ws.protection.sort = False
        ws.protection.set_password(SHEET_PASSWORD)
        ws.protection.enable()
        # print which sheets protected
        print(f"{ws} is protected")


def resize_worksheets(tournament: pd.Series, book: Workbook):
    """Resize the main result/input tables in the OJS workbook to match team count.

    Copies team numbers into each sheet, extends table refs to the required
    number of rows, and replicates formulas, styles, protections and
    conditional formatting from the template row.
    """
    worksheetNames = [
        "Robot Game Scores",
        "Innovation Project Input",
        "Robot Design Input",
        "Core Values Input",
        "Results and Rankings",
    ]
    worksheetTables = [
        "RobotGameScores",
        "InnovationProjectResults",
        "RobotDesignResults",
        "CoreValuesResults",
        "TournamentData",
    ]
    worksheet_start_row = [2, 2, 2, 2, 3]
    div = tournament.get("Div", None)
    # operate on a single open workbook corresponding to the requested division `div`
    if using_divisions:
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments["Short Name"] == tournament["Short Name"])
            & (dfAssignments["Div"] == div)
        ]
    else:
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments["Short Name"] == tournament["Short Name"])
        ]

    # copy the team number data to each of the worksheets
    sheet_tables = zip(worksheetNames, worksheetTables, worksheet_start_row)
    copied_counts: dict[str, int] = {}
    for s, t, r in sheet_tables:
        if s in book.sheetnames:
            ws = book[s]
            tapi_sheet = book["Team and Program Information"]
            # enable debug while testing to see what is read/written
            copied = copy_team_numbers(
                source_sheet=tapi_sheet,
                target_sheet=ws,
                target_start_row=r,
                debug=False,
            )
            copied_counts[s] = copied
            print(f"copy_team_numbers wrote {copied} rows into sheet {s}")

    # Resize the tables
    sheet_tables = zip(worksheetNames, worksheetTables, worksheet_start_row)
    for s, t, r in sheet_tables:
        if s not in book.sheetnames:
            continue
        ws = book[s]
        table: Table = ws.tables[t]
        table_range: str = table.ref

        # parse start/end cells robustly
        start_cell, end_cell = table_range.split(":")
        start_col_letter, start_row_num = coordinate_from_string(start_cell)
        end_col_letter, end_row_num = coordinate_from_string(end_cell)

        # Prefer the actual copied row count for this sheet (fallback to assignees length)
        rows_for_table = copied_counts.get(s, len(assignees))
        # header is at start_row_num; data rows begin at start_row_num+1
        new_end_row = start_row_num + rows_for_table

        # set table.ref correctly using parsed columns
        table.ref = f"{start_col_letter}{start_row_num}:{end_col_letter}{new_end_row}"
        print(
            f"Resized table {t}: new ref={table.ref}, start_row={start_row_num}, rows_for_table={rows_for_table}, assignees={len(assignees)}"
        )

        # Copy formulas: detect any formula in the first data row and copy it down
        first_data_row = start_row_num + 1
        start_col_idx = column_index_from_string(start_col_letter)
        end_col_idx = column_index_from_string(end_col_letter)
        for col_idx in range(
            start_col_idx + 1, end_col_idx + 1
        ):  # skip team number column (A)
            template = ws.cell(row=first_data_row, column=col_idx).value
            if isinstance(template, str) and template.startswith("="):
                for rr in range(first_data_row, new_end_row + 1):
                    ws.cell(row=rr, column=col_idx).value = template

        # Copy cell protection (locked/hidden) from the first data row down through each column
        # This ensures formulas and other cells keep the same locked/hidden state as in the template row.
        for col_idx in range(start_col_idx, end_col_idx + 1):
            template_cell = ws.cell(row=first_data_row, column=col_idx)
            tpl_prot = getattr(template_cell, "protection", None)
            if tpl_prot is None:
                continue
            for rr in range(first_data_row, new_end_row + 1):
                tgt = ws.cell(row=rr, column=col_idx)
                try:
                    # copy the Protection object to the target cell (shallow copy)
                    tgt.protection = _copy(tpl_prot)
                except Exception:
                    # best-effort; don't abort resizing if protection copy fails
                    pass

        # Copy cell style (font, fill, number_format, alignment, border and internal _style)
        # Use shallow copies to avoid sharing mutable style objects between cells.
        for col_idx in range(start_col_idx, end_col_idx + 1):
            template_cell = ws.cell(row=first_data_row, column=col_idx)
            for rr in range(first_data_row, new_end_row + 1):
                tgt = ws.cell(row=rr, column=col_idx)
                try:
                    # copy internal style and common style attributes
                    if hasattr(template_cell, "_style"):
                        tgt._style = _copy(template_cell._style)
                    tgt.number_format = template_cell.number_format
                    tgt.font = _copy(template_cell.font)
                    tgt.fill = _copy(template_cell.fill)
                    tgt.alignment = _copy(template_cell.alignment)
                    tgt.border = _copy(template_cell.border)
                except Exception:
                    # best-effort: keep going if any individual copy fails
                    pass

        # Copy row height from template row to newly created rows (if a height is set)
        try:
            template_height = ws.row_dimensions[first_data_row].height
            if template_height is not None:
                for rr in range(first_data_row, new_end_row + 1):
                    ws.row_dimensions[rr].height = template_height
        except Exception:
            pass

        # remove any rows below the new end row to keep sheet tidy
        # Before deleting extra rows, replicate data-validation and conditional-formatting
        try:
            _copy_data_validations_for_range(
                ws, start_col_idx, end_col_idx, first_data_row, new_end_row
            )
        except Exception:
            pass

        try:
            _extend_conditional_formatting_for_range(
                ws, start_col_letter, end_col_letter, first_data_row, new_end_row
            )
        except Exception:
            pass

        ws.delete_rows(idx=new_end_row + 1, amount=200)
        # ws.protection.sheet = True

    # Do not save here; caller should save/close the workbook once after calling helpers.


def copy_award_def(tournament: pd.Series, book: Workbook):
    # Add AwardDef table to the open workbook `book`
    add_table_dataframe(book, "AwardDef", "AwardDef", dfAwardDef)


def add_conditional_formats(tournament: pd.Series, book: Workbook):
    # Create fill
    greenAwardFill = PatternFill(
        start_color="00B050", end_color="00B050", fill_type="solid"
    )
    greenAdvFill = PatternFill(
        start_color="00FF00", end_color="00FF00", fill_type="solid"
    )
    rgGoldFill = PatternFill(
        start_color="C9B037", end_color="C9B037", fill_type="solid"
    )
    rgSilverFill = PatternFill(
        start_color="D7D7D7", end_color="D7D7D7", fill_type="solid"
    )
    rgBronzeFill = PatternFill(
        start_color="AD8A56", end_color="AD8A56", fill_type="solid"
    )
    # operate on the provided open workbook `book`
    ws = book["Results and Rankings"]
    # Award
    ws.conditional_formatting.add(
        "O2",
        FormulaRule(
            formula=["COUNTA(AwardList!$A$2:$A$35)=COUNTA($O$3:$O$288)"],
            stopIfTrue=False,
            fill=greenAwardFill,
        ),
    )
    # Advancing
    ws.conditional_formatting.add(
        "P2",
        FormulaRule(
            formula=['COUNTIF($P:$P,"Yes")=Meta!$B$13'],
            stopIfTrue=False,
            fill=greenAdvFill,
        ),
    )
    # Robot Game Gold
    ws.conditional_formatting.add(
        "J1:J100",
        FormulaRule(
            formula=[
                'AND(J1=1,IF(VLOOKUP("Robot Game 1st Place",AwardList!$C$2:$C$7,1,FALSE)="Robot Game 1st Place", TRUE, FALSE))'
            ],
            stopIfTrue=False,
            fill=rgGoldFill,
        ),
    )
    # Robot Game Silver
    ws.conditional_formatting.add(
        "J1:J100",
        FormulaRule(
            formula=[
                'AND(J1=2,IF(VLOOKUP("Robot Game 2nd Place",AwardList!$C$2:$C$7,1,FALSE)="Robot Game 2nd Place", TRUE, FALSE))'
            ],
            stopIfTrue=False,
            fill=rgSilverFill,
        ),
    )
    # Robot Game Bronze
    ws.conditional_formatting.add(
        "J1:J100",
        FormulaRule(
            formula=[
                'AND(J1=3,IF(VLOOKUP("Robot Game 3rd Place",AwardList!$C$2:$C$7,1,FALSE)="Robot Game 3rd Place", TRUE, FALSE))'
            ],
            stopIfTrue=False,
            fill=rgBronzeFill,
        ),
    )


def hide_worksheets(tournament: pd.Series, book: Workbook):
    worksheetNames = ["Data Validation", "Meta", "AwardListDropdowns", "AwardDef"]
    # operate on the provided open workbook `book`
    for sheetname in worksheetNames:
        if sheetname in book.sheetnames:
            ws = book[sheetname]
            ws.sheet_state = "hidden"


# # # # # # # # # # # # # # # # # # # # #

# If this line isn't here, we will get a UserWarning when we run the program, alerting us that the
# conditional formatting in the spreadsheets will not be preserved when we copy the data.
# We don't need conditional formatting within the data manipulation, so it isn't a big deal
# UserWarning: Data Validation extension is not supported and will be removed
# https://stackoverflow.com/questions/53965596/python-3-openpyxl-userwarning-data-validation-extension-not-supported
warnings.simplefilter(action="ignore", category=UserWarning)

# Initialize colorama
init()

cwd: str = os.getcwd()
if getattr(sys, "frozen", False):
    dir_path = os.path.dirname(sys.executable)
elif __file__:
    dir_path = os.path.dirname(__file__)

config = load_json_without_notes(dir_path + "\\" + "season.json")

tournament_file: str = dir_path + "\\" + config["filename"]
template_file: str = dir_path + "\\" + config["tournament_template"]
extrafilelist: list[str] = config["copy_file_list"]

# Make sure the extra files exist
for filename in extrafilelist:
    try:
        if os.path.exists(dir_path + "\\" + filename):
            print(
                Fore.GREEN + f"{filename}... CHECK!",
            )
        else:
            print(Fore.RED)
            print(f"{dir_path + "\\" + filename}... MISSING!")
    except Exception as e:
        print_error(f"Got an error checking for {filename}", e)

# Read the state tournament workbook to get the details for all of the events
try:
    print(Fore.RESET)
    print(f"Getting tournaments from {dir_path}")
    book = load_workbook(tournament_file, data_only=True)
    book.close()
except Exception as e:
    print_error(
        f"Could not open the tournament file: {tournament_file}. Check to make sure it is not open in Excel.",
        e,
    )

try:
    dictSeasonInfo: pd.DataFrame = read_table_as_dict(
        tournament_file, "SeasonInfo", "SeasonInfo"
    )
    using_divisions: bool = dictSeasonInfo["Divisions"]
except Exception as e:
    print_error(f"Could not read the SeasonInfo table.", e)

try:
    if using_divisions:
        dfTournaments = read_table_as_df(
            tournament_file, "DivTournaments", "DivTournamentList"
        ).fillna(0)
    else:
        dfTournaments = read_table_as_df(
            tournament_file, "Tournaments", "TournamentList"
        ).fillna(0)
except Exception as e:
    print_error(f"Could not read the tournament worksheet.", e)

try:
    dfAwardDef = read_table_as_df(tournament_file, "AwardDef", "AwardDef").fillna(0)
except Exception as e:
    print_error(f"Could not read the AwardDef worksheet.", e)

try:
    dfAssignments = read_table_as_df(
        tournament_file, "Assignments", "Assignments"
    ).fillna(0)
    tourn_array: list[str] = []
    for index, row in dfTournaments.iterrows():
        tourn_array.append(row["Short Name"])
    print("Assignments")
    print(dfAssignments)
except Exception as e:
    print_error(f"Could not read the assignments worksheet.", e)


# Are we building all of the tournaments, or just one?
print(Fore.RESET)
tourn = input("Enter the tournament short name, or press ENTER for all tournaments: ")
if tourn != "":
    if tourn in tourn_array:
        dfTournaments = dfTournaments.loc[dfTournaments["Short Name"] == tourn]
    else:
        input(
            f"Tournament not found. The tournament name must come from this list: {tourn_array}\nPress enter to exit..."
        )
        sys.exit(1)

print(dfTournaments)

# Now that we have all of the info for the tournaments, loop through and
# start building the OJS files and folders
for index, row in dfTournaments.iterrows():
    newpath = dir_path + "\\tournaments\\" + row["Short Name"]
    create_folder(newpath)
    print("Copying files")
    copy_files(row)

    # For divisions, there may be separate OJS files for each division (D1/D2).
    ojs_name = row.get("OJS_FileName")
    if ojs_name is None or (isinstance(ojs_name, float) and pd.isna(ojs_name)):
        print(f"Did not see {ojs_name}")
        continue
    ojs_path = dir_path + "\\tournaments\\" + row["Short Name"] + "\\" + ojs_name
    print(ojs_path)
    ojs_book = load_workbook(ojs_path, read_only=False, keep_vba=True)
    try:
        set_up_tapi_worksheet(row, ojs_book)
        set_up_award_worksheet(row, ojs_book)
        set_up_meta_worksheet(row, ojs_book)
        add_conditional_formats(row, ojs_book)
        copy_award_def(row, ojs_book)
        hide_worksheets(row, ojs_book)
        resize_worksheets(row, ojs_book)
        protect_worksheets(row, ojs_book)
    finally:
        ojs_book.save(ojs_path)
        ojs_book.close()
