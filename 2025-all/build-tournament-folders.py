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

# pip install pandas
import pandas as pd

# pip install xlwings
import xlwings

import math
import numpy as np

# Constants - Configuration
SHEET_PASSWORD: str = "skip"
CONFIG_FILENAME: str = "season.json"

# Constants - Column names
COL_TEAM_NUMBER: str = "Team #"
COL_TEAM_NAME: str = "Team Name"
COL_COACH_NAME: str = "Coach Name"
COL_POD_NUMBER: str = "Pod Number"
COL_SHORT_NAME: str = "Short Name"
COL_LONG_NAME: str = "Long Name"
COL_OJS_FILENAME: str = "OJS_FileName"
COL_DIVISION: str = "Div"
COL_ADVANCING: str = "ADV"

# Constants - Sheet names
SHEET_TEAM_INFO: str = "Team and Program Information"
SHEET_AWARD_DROPDOWNS: str = "AwardListDropdowns"
SHEET_META: str = "Meta"
SHEET_AWARD_DEF: str = "AwardDef"
SHEET_ROBOT_GAME: str = "Robot Game Scores"
SHEET_INNOVATION: str = "Innovation Project Input"
SHEET_ROBOT_DESIGN: str = "Robot Design Input"
SHEET_CORE_VALUES: str = "Core Values Input"
SHEET_RESULTS: str = "Results and Rankings"

# Constants - Table names
TABLE_TEAM_LIST: str = "OfficialTeamList"
TABLE_ROBOT_GAME_AWARDS: str = "RobotGameAwards"
TABLE_AWARD_DROPDOWNS: str = "AwardListDropdowns"
TABLE_META: str = "Meta"
TABLE_AWARD_DEF: str = "AwardDef"
TABLE_ROBOT_GAME: str = "RobotGameScores"
TABLE_INNOVATION: str = "InnovationProjectResults"
TABLE_ROBOT_DESIGN: str = "RobotDesignResults"
TABLE_CORE_VALUES: str = "CoreValuesResults"
TABLE_TOURNAMENT_DATA: str = "TournamentData"

# Constants - Folder structure
FOLDER_TOURNAMENTS: str = "tournaments"
FILE_CLOSING_CEREMONY: str = "closing_ceremony.html"

# Constants - Award columns
AWARD_COLUMN_PREFIX_JUDGED: str = "J_"
AWARD_COLUMN_ROBOT_GAME: str = "P_AWD_RG"
AWARD_LABEL_PREFIX: str = "Label"

import json
from typing import Union, Dict, Any, Sequence


def print_error(errormsg: str, e: Exception | None = None) -> None:
    """Print an error message in red and exit the program.
    
    Args:
        errormsg: The error message to display
        e: Optional exception object to include in the output
    """
    print(Fore.RED)
    if e:
        print(f"{errormsg}\n{e}")
    else:
        print(errormsg)
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
    debug: bool = False,
) -> int:
    """
    Append a pandas.DataFrame to an existing Excel table.
    
    Args:
        wb: The Excel workbook object
        sheet_name: Name of the worksheet containing the table
        table_name: Name of the Excel table to append to
        data: DataFrame containing the data to append
        require_all_columns: If True, validate that DataFrame columns match table headers exactly
        keep_vba: Placeholder for VBA preservation (not currently used)
        debug: If True, print diagnostic information
        
    Returns:
        Number of rows written to the table
        
    Raises:
        KeyError: If sheet or table is not found
        ValueError: If DataFrame columns don't match table headers (when require_all_columns=True)
    """
    if data is None or data.empty:
        print_error("Attempting to add empty or None DataFrame to table. "
                   f"Sheet: {sheet_name}, Table: {table_name}")
        return 0

    if sheet_name not in wb.sheetnames:
        print_error(f"Sheet '{sheet_name}' not found in workbook. "
                   f"Available sheets: {', '.join(wb.sheetnames)}")
    ws = wb[sheet_name]

    if table_name not in ws.tables:
        available_tables = ', '.join(ws.tables.keys()) if ws.tables else 'none'
        print_error(f"Table '{table_name}' not found on sheet '{sheet_name}'. "
                   f"Available tables: {available_tables}")
    
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

    # Debug: print table/dataframe diagnostics
    if debug:
        print(Fore.RESET)
        print(f"add_table_dataframe: sheet={sheet_name}, table={table_name}")
        print(f"DataFrame shape: rows={len(df)}, cols={len(df.columns)}")
        print(f"Columns: {list(df.columns)}")
        print(f"Table range: {table_range} (start={start_cell}, end={end_cell})")
        try:
            end_col_letter = coordinate_from_string(end_cell)[0]
            end_col_idx = column_index_from_string(end_col_letter)
            print(f"Start col idx: {start_col_idx}, End col idx: {end_col_idx}, header row: {start_row}, data rows end: {end_row}")
        except Exception:
            pass
        # show a small preview
        try:
            print("DataFrame preview:")
            print(df.head(8))
        except Exception:
            pass

    rows_written = 0
    df_iter_index = 0
    total_rows = len(df)

    filled_blank_rows = 0
    appended_count = 0

    # Fill existing blank data rows first
    for row_tuple in table_data_rows:
        if df_iter_index >= total_rows:
            if debug:
                print("add_table_dataframe: no more rows to write; stopping fill of blank rows")
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
            filled_blank_rows += 1

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
        appended_count += 1

    # If we appended rows, extend the table ref to include the new end row
    if appended:
        end_col_letter = coordinate_from_string(end_cell)[0]
        table.ref = f"{start_cell}:{end_col_letter}{current_row}"
        if debug:
            print(f"add_table_dataframe: appended {appended_count} rows, new table.ref={table.ref}")

    if debug:
        print(f"add_table_dataframe: filled_blank_rows={filled_blank_rows}, appended_count={appended_count}, total_rows_written={rows_written}")

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

    Args:
        xlsx_path: Path to the Excel workbook
        sheet_name: Name of the worksheet containing the table
        table_name: Name of the Excel table to read
        require_table: If True, raise an error if table is not found
        convert_integer_floats: If True, convert float columns with whole numbers to Int64
        
    Returns:
        DataFrame containing the table data
        
    Raises:
        RuntimeError: If workbook cannot be opened
        KeyError: If sheet or table is not found (when require_table=True)
        ValueError: If table reference format is invalid
    """
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"Excel file not found: {xlsx_path}")
    
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

    Args:
        xlsx_path: Path to the Excel workbook
        sheet_name: Name of the worksheet containing the table
        table_name: Name of the Excel table to read
        key_col: Column name to use as key (defaults to first column)
        value_col: Column name to use as value (defaults to second column)
        require_unique_keys: If True, raise error on duplicate keys
        
    Returns:
        Dictionary mapping keys to values
        
    Raises:
        ValueError: If table doesn't have exactly 2 columns or has duplicate keys
        KeyError: If specified key/value columns are not found
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
    
    Args:
        path: Path to the JSON file
        
    Returns:
        Dictionary with note keys removed
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        json.JSONDecodeError: If the file contains invalid JSON
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"Configuration file not found: {path}")
    
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError(
            f"Invalid JSON in configuration file: {path}",
            e.doc,
            e.pos
        ) from e
    
    return _remove_note_keys(data)


def create_folder(newpath: str) -> None:
    """Create `newpath` if it does not exist.

    Args:
        newpath: Path to the directory to create
        
    Raises:
        Exits via print_error if directory creation fails
    """
    if not newpath or not newpath.strip():
        print_error("Cannot create folder: path is empty or None")
        
    if os.path.exists(newpath):
        return
    
    try:
        os.makedirs(newpath)
        print(Fore.RESET)
        print(f"Created folder: {newpath}")
    except PermissionError as e:
        print_error(f"Permission denied when creating directory: {newpath}", e)
    except OSError as e:
        print_error(f"OS error when creating directory: {newpath}", e)
    except Exception as e:
        print_error(f"Unexpected error creating directory: {newpath}", e)


def copy_files(item: pd.Series) -> None:
    """Copy extra files and the OJS template into the tournament folder.

    Args:
        item: A pandas Series representing a tournament row with 'Short Name' and 'OJS_FileName'
        
    Global variables used:
        extrafilelist: List of filenames to copy
        template_file: Path to the OJS template workbook
        dir_path: Base directory containing tournaments
        
    Raises:
        Exits via print_error if any copy operation fails
    """
    if COL_SHORT_NAME not in item or not item[COL_SHORT_NAME]:
        print_error(f"Tournament row missing required field '{COL_SHORT_NAME}'")
        
    if COL_OJS_FILENAME not in item or not item[COL_OJS_FILENAME]:
        print_error(f"Tournament row missing required field '{COL_OJS_FILENAME}' for {item[COL_SHORT_NAME]}")
    
    for filename in extrafilelist:
        try:
            source_path = os.path.join(dir_path, filename)
            if not os.path.exists(source_path):
                print_error(f"Source file not found: {source_path}")
                
            newpath = os.path.join(dir_path, FOLDER_TOURNAMENTS, item[COL_SHORT_NAME])
            shutil.copy(source_path, newpath)
        except PermissionError as e:
            print_error(f"Permission denied copying file '{filename}' to {newpath}", e)
        except Exception as e:
            print_error(f"Could not copy file '{filename}' to {newpath}", e)
            
    print("Files copied successfully")
    
    try:
        if not os.path.exists(template_file):
            print_error(f"OJS template file not found: {template_file}")
            
        new_ojs_file = os.path.join(
            dir_path,
            FOLDER_TOURNAMENTS,
            item[COL_SHORT_NAME],
            item[COL_OJS_FILENAME]
        )
        shutil.copy(template_file, new_ojs_file)
    except Exception as e:
        print_error(f"Could not copy OJS template '{template_file}' to '{new_ojs_file}'", e)
        
    print("OJS file(s) copied successfully")


def set_up_tapi_worksheet(tournament: pd.Series, book: Workbook) -> None:
    """Populate the 'Team and Program Information' table in the open workbook.

    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object for the tournament's OJS file
        
    Global variables used:
        dfAssignments: DataFrame containing team assignments
        using_divisions: Boolean indicating if divisions are used
    """
    d = ""
    print(Fore.RESET)
    
    if using_divisions:
        d = tournament[COL_DIVISION]
        if isinstance(tournament[COL_OJS_FILENAME], float):
            print_error(f"Invalid OJS filename for tournament {tournament[COL_SHORT_NAME]}: "
                       "expected string, got float (possibly missing value)")
        print(
            f"Setting up Tournament Team and Program Information for {tournament[COL_SHORT_NAME]} {d}"
        )
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments[COL_SHORT_NAME] == tournament[COL_SHORT_NAME])
            & (dfAssignments[COL_DIVISION] == d)
        ]
        print(
            f'There are {len(assignees)} teams in this {tournament[COL_SHORT_NAME]} {d} tournament'
        )
    else:
        if isinstance(tournament[COL_OJS_FILENAME], float):
            print_error(f"Invalid OJS filename for tournament {tournament[COL_SHORT_NAME]}: "
                       "expected string, got float (possibly missing value)")
        print(
            f"Setting up Tournament Team and Program Information for {tournament[COL_SHORT_NAME]}"
        )
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments[COL_SHORT_NAME] == tournament[COL_SHORT_NAME])
        ]
        print(
            f'There are {len(assignees)} teams in this {tournament[COL_SHORT_NAME]} tournament'
        )

    # Validate required columns exist
    keep = [COL_TEAM_NUMBER, COL_TEAM_NAME, COL_COACH_NAME]
    keep_safe = [c for c in keep if c in assignees.columns]
    
    if len(keep_safe) != len(keep):
        missing = set(keep) - set(keep_safe)
        print_error(f"Missing required columns in assignments: {', '.join(missing)}")
    
    assignees = assignees[keep_safe]
    assignees[COL_POD_NUMBER] = 0
    sorted_assignees = assignees.sort_values(by=COL_TEAM_NUMBER, ascending=True)
    print(sorted_assignees)
    
    add_table_dataframe(
        book,
        SHEET_TEAM_INFO,
        TABLE_TEAM_LIST,
        sorted_assignees
    )


def set_up_award_worksheet(tournament: pd.Series, book: Workbook) -> None:
    """Prepare award tables used by OJS dropdowns and closing scripts.

    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        
    Global variables used:
        dfAwardDef: DataFrame containing award definitions
        using_divisions: Boolean indicating if divisions are used
    """
    thisDiv = tournament[COL_DIVISION] if using_divisions else ""
    print(Fore.RESET)
    print(
        f"Setting up Tournament Division Awards for {tournament[COL_SHORT_NAME]} {thisDiv}"
    )

    # Robot Game awards
    rg_raw = (
        tournament.get(AWARD_COLUMN_ROBOT_GAME)
        if hasattr(tournament, "get")
        else tournament[AWARD_COLUMN_ROBOT_GAME]
    )
    try:
        rg_awards = int(rg_raw) if not pd.isna(rg_raw) else 0
    except (ValueError, TypeError) as e:
        print_error(f"Invalid robot game award count for {tournament[COL_SHORT_NAME]}: {rg_raw}", e)
        
    rg_awards_df = pd.DataFrame(columns=["Robot Game Awards"])
    for rg_award_num in range(1, rg_awards + 1):
        thisLabel = AWARD_LABEL_PREFIX + str(rg_award_num)
        sel = dfAwardDef.loc[dfAwardDef["ColumnName"] == AWARD_COLUMN_ROBOT_GAME, thisLabel]
        try:
            thisValue = sel.iat[0]
        except (IndexError, KeyError):
            thisValue = None
        rg_awards_df.loc[len(rg_awards_df)] = [thisValue]

    add_table_dataframe(book, SHEET_AWARD_DROPDOWNS, TABLE_ROBOT_GAME_AWARDS, rg_awards_df)

    # Other judged awards
    j_cols = tournament.filter(regex=f"^{AWARD_COLUMN_PREFIX_JUDGED}")
    j_awards_df = pd.DataFrame(columns=["Award"])
    
    for this_col_name, series in j_cols.items():
        count = _to_int(series)
        for award_num in range(1, count + 1):
            label_col = AWARD_LABEL_PREFIX + str(award_num)
            sel = dfAwardDef.loc[dfAwardDef["ColumnName"] == this_col_name, label_col]
            try:
                thisValue = sel.iat[0]
            except (IndexError, KeyError):
                thisValue = None
            j_awards_df.loc[len(j_awards_df)] = [thisValue]
            
    add_table_dataframe(book, SHEET_AWARD_DROPDOWNS, TABLE_AWARD_DROPDOWNS, j_awards_df)


def set_up_meta_worksheet(tournament: pd.Series, book: Workbook) -> None:
    """Populate the metadata worksheet with tournament information.
    
    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        
    Global variables used:
        config: Dictionary containing season configuration
        dir_path: Base directory path
        using_divisions: Boolean indicating if divisions are used
    """
    print(Fore.RESET)
    d = tournament[COL_DIVISION] if using_divisions else ""
    print(f"Setting up meta worksheet for {tournament[COL_SHORT_NAME]} {d}")
    
    scriptfile = os.path.join(
        dir_path,
        FOLDER_TOURNAMENTS,
        tournament[COL_SHORT_NAME],
        FILE_CLOSING_CEREMONY
    )
    
    # Validate required config keys
    required_keys = ["season_yr", "season_name"]
    for key in required_keys:
        if key not in config:
            print_error(f"Missing required configuration key: '{key}' in {CONFIG_FILENAME}")
    
    dfMeta = pd.DataFrame(columns=["Key", "Value"])
    dfMeta.loc[len(dfMeta)] = {"Key": "Tournament Year", "Value": config["season_yr"]}
    dfMeta.loc[len(dfMeta)] = {"Key": "FLL Season Title", "Value": config["season_name"]}
    dfMeta.loc[len(dfMeta)] = {"Key": "Tournament Short Name", "Value": tournament[COL_SHORT_NAME]}
    dfMeta.loc[len(dfMeta)] = {"Key": "Tournament Long Name", "Value": tournament[COL_LONG_NAME]}
    dfMeta.loc[len(dfMeta)] = {"Key": "Completed Script File", "Value": scriptfile}
    dfMeta.loc[len(dfMeta)] = {"Key": "Using Divisions", "Value": using_divisions}
    
    if using_divisions:
        dfMeta.loc[len(dfMeta)] = {"Key": "Division", "Value": d}
        
    dfMeta.loc[len(dfMeta)] = {"Key": "Advancing", "Value": tournament[COL_ADVANCING]}

    print('Adding meta table to the OJS')
    
    if tournament[COL_OJS_FILENAME] is not None:
        add_table_dataframe(book, SHEET_META, TABLE_META, dfMeta, debug=False)


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
    print(f"Protecting worksheets for {tournament['OJS_FileName']}")
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
        # print(f"{ws} is protected")
    print("Worksheets protected")


def resize_worksheets(tournament: pd.Series, book: Workbook) -> None:
    """Resize the main result/input tables in the OJS workbook to match team count.

    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        
    Global variables used:
        dfAssignments: DataFrame containing team assignments
        using_divisions: Boolean indicating if divisions are used
    """
    print(f"Resizing worksheets for {tournament[COL_OJS_FILENAME]}")
    
    worksheetNames = [
        SHEET_ROBOT_GAME,
        SHEET_INNOVATION,
        SHEET_ROBOT_DESIGN,
        SHEET_CORE_VALUES,
        SHEET_RESULTS,
    ]
    worksheetTables = [
        TABLE_ROBOT_GAME,
        TABLE_INNOVATION,
        TABLE_ROBOT_DESIGN,
        TABLE_CORE_VALUES,
        TABLE_TOURNAMENT_DATA,
    ]
    worksheet_start_row = [2, 2, 2, 2, 3]
    
    div = tournament.get(COL_DIVISION, None)
    
    if using_divisions:
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments[COL_SHORT_NAME] == tournament[COL_SHORT_NAME])
            & (dfAssignments[COL_DIVISION] == div)
        ]
    else:
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments[COL_SHORT_NAME] == tournament[COL_SHORT_NAME])
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
            # print(f"copy_team_numbers wrote {copied} rows into sheet {s}")

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
        # print(
        #     f"Resized table {t}: new ref={table.ref}, start_row={start_row_num}, rows_for_table={rows_for_table}, assignees={len(assignees)}"
        # )

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


def copy_award_def(tournament: pd.Series, book: Workbook) -> None:
    """Add the award definitions table to the workbook.
    
    Args:
        tournament: A pandas Series representing the tournament row (for context)
        book: An open openpyxl Workbook object
        
    Global variables used:
        dfAwardDef: DataFrame containing award definitions
    """
    add_table_dataframe(book, SHEET_AWARD_DEF, TABLE_AWARD_DEF, dfAwardDef)


def add_conditional_formats(tournament: pd.Series, book: Workbook) -> None:
    """Add conditional formatting rules to the Results and Rankings worksheet.
    
    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
    """
    print(f"Adding conditional formats for {tournament[COL_OJS_FILENAME]}")
    
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


def hide_worksheets(tournament: pd.Series, book: Workbook) -> None:
    """Hide utility worksheets that should not be visible to users.
    
    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
    """
    print(f"Hiding worksheets for {tournament[COL_OJS_FILENAME]}")

    worksheetNames = ["Data Validation", SHEET_META, SHEET_AWARD_DROPDOWNS, SHEET_AWARD_DEF]
    
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

try:
    config = load_json_without_notes(os.path.join(dir_path, CONFIG_FILENAME))
except FileNotFoundError:
    print_error(f"Configuration file '{CONFIG_FILENAME}' not found in {dir_path}")
except json.JSONDecodeError as e:
    print_error(f"Invalid JSON in configuration file '{CONFIG_FILENAME}'", e)

# Validate required config keys
required_config_keys = ["filename", "tournament_template", "copy_file_list"]
missing_keys = [key for key in required_config_keys if key not in config]
if missing_keys:
    print_error(f"Missing required keys in {CONFIG_FILENAME}: {', '.join(missing_keys)}")

tournament_file: str = os.path.join(dir_path, config["filename"])
template_file: str = os.path.join(dir_path, config["tournament_template"])
extrafilelist: list[str] = config["copy_file_list"]

# Validate that required files exist
if not os.path.exists(tournament_file):
    print_error(f"Tournament file not found: {tournament_file}")
    
if not os.path.exists(template_file):
    print_error(f"Template file not found: {template_file}")

# Make sure the extra files exist
for filename in extrafilelist:
    try:
        if os.path.exists(os.path.join(dir_path, filename)):
            print(
                Fore.GREEN + f"{filename}... CHECK!",
            )
        else:
            print(Fore.RED)
            print(f"{os.path.join(dir_path, filename)}... MISSING!")
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
    print(Fore.YELLOW)
    print("* * * * * * * * * * * * * * * * * * * * * *")
    print(str(str(row["Short Name"]) + " " + str(row.get("Div", ""))).center(43))
    print("* * * * * * * * * * * * * * * * * * * * * *")
    print(Fore.RESET)
    newpath = os.path.join(dir_path, "tournaments", row["Short Name"])
    create_folder(newpath)
    print(row.to_string())
    print("----------------------------------------")
    print("Copying files")
    copy_files(row)

    # For divisions, there may be separate OJS files for each division (D1/D2).
    ojs_name = row.get("OJS_FileName")
    if ojs_name is None or (isinstance(ojs_name, float) and pd.isna(ojs_name)):
        print(f"Did not see {ojs_name}")
        continue
    ojs_path = os.path.join(dir_path, "tournaments", row["Short Name"], ojs_name)
    print(f"ojs_path: {ojs_path}")
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

    print(Fore.GREEN)
    print(f"Completed setup for {row['Short Name']} {str(row.get("Div", ""))}")
    print(Fore.RESET)

print(Fore.GREEN)
print(f"All done!")
print(Fore.RESET)