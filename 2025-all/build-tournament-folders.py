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

def add_table_dataframe(
    xlsx_path: str,
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
        return 0

    wb = load_workbook(xlsx_path, read_only=False, keep_vba=keep_vba)
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]

    if table_name not in ws.tables:
        raise KeyError(f"Table {table_name!r} not found on sheet {sheet_name!r}")
    table = ws.tables[table_name]
    table_range = table.ref  # e.g. "C2:F20"

    # header row and data rows within the table ref
    table_head = ws[table_range][0]
    table_data_rows = ws[table_range][1:]
    headers = [c.value.strip() if isinstance(c.value, str) else c.value for c in table_head]

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
            break
        # consider a row blank if all cells are None or blank strings
        if all(
            (cell.value is None) or (isinstance(cell.value, str) and cell.value.strip() == "")
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

    wb.save(xlsx_path)
    return rows_written


def read_table_as_df(xlsx_path: str, sheet_name: str, table_name: str, require_table: bool = True) -> pd.DataFrame:
    """
    Read an Excel table (ListObject) by name into a pandas DataFrame.

    Safer behaviour:
    - Catches workbook / sheet / table access errors and re-raises with context.
    - Validates the table.ref format.
    - Handles empty tables (returns empty DataFrame when require_table==False).
    - Cleans column names and trims string values using vectorized operations.
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
            raise KeyError(f"Table {table_name!r} not found on sheet {sheet_name!r} in {xlsx_path}")
        return pd.DataFrame()

    table = ws.tables[table_name]
    ref = table.ref  # expected e.g. "C2:F20"
    if not isinstance(ref, str) or ":" not in ref:
        raise ValueError(f"Unexpected table.ref for {table_name!r} on {sheet_name!r}: {ref!r}")

    try:
        start, end = ref.split(":")
        start_col, start_row = coordinate_from_string(start)
        end_col, end_row = coordinate_from_string(end)
    except Exception as e:
        raise ValueError(f"Could not parse table.ref '{ref}' for {table_name!r}: {e}") from e

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
        raise RuntimeError(f"pandas.read_excel failed for table {table_name!r} on sheet {sheet_name!r}: {e}") from e

    # basic cleanup: strip column names and trim string values safely (vectorized where possible)
    df.columns = df.columns.str.strip()

    def _trim_series(s: pd.Series) -> pd.Series:
        if pd.api.types.is_string_dtype(s):
            return s.str.strip()
        if s.dtype == object:
            return s.map(lambda v: v.strip() if isinstance(v, str) else v)
        return s

    df = df.apply(_trim_series)

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
        raise ValueError(f"Table {table_name!r} on sheet {sheet_name!r} must have exactly 2 columns (found {df.shape[1]})")

    # Determine which columns to use
    col_names = list(df.columns)
    key_col_name = key_col if key_col is not None else col_names[0]
    value_col_name = value_col if value_col is not None else col_names[1]

    if key_col_name not in df.columns or value_col_name not in df.columns:
        raise KeyError(f"Specified key/value columns not found in table columns: {df.columns.tolist()}")

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
            raise ValueError(f"Duplicate key found in table {table_name!r} on sheet {sheet_name!r}: {key!r}")
        mapping[key] = val

    return mapping

def _remove_note_keys(obj: Any) -> Any:
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
    print(item)
    for filename in extrafilelist:
        try:
            newpath = dir_path + "\\tournaments\\" + item["Short Name"] + "\\"
            shutil.copy(dir_path + "\\" + filename, newpath)
        except Exception as e:
            print_error(f"Could not copy file: {filename} to {newpath}", e)
    # next copy the template files
    ojsfilelist = []
    if pd.notna(item["D1_OJS"]):
        ojsfilelist.append(item["D1_OJS"])
    if pd.notna(item["D2_OJS"]):
        ojsfilelist.append(item["D2_OJS"])

    for filename in ojsfilelist:
        try:
            folder = dir_path + "\\tournaments\\" + item["Short Name"] + "\\"
            print(type(folder), folder)
            print(type(filename), filename)
            shutil.copy(
                template_file,
                folder + filename,
            )
        except Exception as e:
            print_error(f'Could not copy OJS file: {template_file} to {folder}\\{filename}', e)
    
    # create a file list for the tournament folder
    directory = dir_path + "\\tournaments\\" + item["Short Name"]
    file_list = os.listdir(directory)
    # Remove any files we don't want in the file_list
    # in particular, we won't need the script_maker programs for two reasons
    # 1) the script_maker program is the program they will be running, so it can't be missing
    # 2) no need to check if the script_maker for other OS'es are present
    try:
        file_list.remove("script_maker-win.exe")
    except:
        pass

    try:
        file_list.remove("script_maker-mac")
    except:
        pass


# edits the OJS spreadsheets with the correct tournament information
def set_up_tapi_worksheet(tournament: pd.Series):
    # open the OJS workbook
    for d in ["D1", "D2"]:
        # print(type(tournament[d + "_OJS"]), tournament[d + "_OJS"])
        if isinstance(tournament[d + "_OJS"], float):
            continue
        print(f"Setting up Tournament Team and Program Information for {tournament["Short Name"]} {d}")
        divassignees: pd.DataFrame = dfAssignments[
            (dfAssignments["Tournament"] == tournament["Short Name"])
            & (dfAssignments["Div"] == d)
        ]
        print(divassignees)
        print(Fore.RESET)
        print(f'There are {len(divassignees)} teams in this {d} {tournament["Short Name"]} tournament')
        if len(divassignees.index) > 0:
            try:
                # print(divassignees)
                ojsfile = (
                    dir_path
                    + "\\tournaments\\"
                    + tournament["Short Name"]
                    + "\\"
                    + tournament[d + "_OJS"]
                )
                ojs_book = load_workbook(ojsfile, read_only=False, keep_vba=True)
                ws = ojs_book["Team and Program Information"]
                table: Table = ws.tables["OfficialTeamList"]
                table_range: str = table.ref  # should return a string C2:F3
                start_cell = table_range.split(":")[0]  # should return 'C2'
                coords = coordinate_from_string(start_cell)
                start_col = column_index_from_string(coords[0])
                start_index = divassignees.index[0]
                colonPos = (table.ref).find(":")
                print(Fore.RESET)
                print(f'About to resize the table {table.ref}, {table.ref[:colonPos]}, {re.sub(r'\d', '', table.ref[colonPos + 1])}, {str(len(divassignees) + 2)}')
                table.ref = (
                    table.ref[:colonPos]
                    + ":"
                    + re.sub(r"\d", "", table.ref[colonPos + 1])
                    + str(len(divassignees) + 2)
                )
                for i, row in divassignees.iterrows():
                    # cell = f'{get_column_letter(start_col)}{i + 3 - start_index}'
                    # print(f'Cell: {cell}, setting value {row['Team #']}')
                    # ws[cell] = row['Team #']
                    thisrow = i + 3 - start_index
                    ws.cell(row=thisrow, column=start_col).value = row["Team #"]
                    ws.cell(row=thisrow, column=start_col + 1).value = row["Team Name"]
                    ws.cell(row=thisrow, column=start_col + 2).value = row["Coach Name"]

                # Save the workbook
                print(Fore.RESET)
                print(f"Saving workbook. OfficialTeamList ref: {table.ref}")
                ojs_book.save(ojsfile)
                # ojs_book.close()
            except Exception as e:
                print_error(f"There was an error setting up the TAPI worksheet:", e)

def set_up_award_worksheet(tournament: pd.Series):
    # This will read the awards allocations for each tournament (at the tournament
    # level--no divisions) and create the AwardList table which goes on the
    # normally hidden AwardList worksheet. This list is used to create the
    # dropdowns in the award column on the OJS
    s1 = tournament.fillna(0)
    ojsfile = (
        dir_path
        + "\\tournaments\\"
        + s1["Short Name"]
        + "\\"
        + s1["D1_OJS"]
    )
    print(Fore.RESET)
    print(f"Setting up Tournament Awards for {s1["Short Name"]}")
    j_cols = s1.filter(regex=f"^J_")
    print(j_cols)
    j_awards_df = pd.DataFrame(columns=["Award"])
    for this_col_name, series in j_cols.items():
        print(f"Working the {this_col_name} awards.")
        print(f"There are {series} places for this award.")
        for award_num in range(0, int(series)):
            # print("Adding new judged award")
            # print(find_item_by_col_name(d=config, col_name=this_col_name))
            # j_awards_df.loc[len(j_awards_df)] = [config["TOURNAMENT_AWARDS"][find_item_by_col_name(d=config["TOURNAMENT_AWARDS"], col_name=this_col_name)]["dropdown"]["p" + str(award_num + 1)]]
            j_awards_df.loc[len(j_awards_df)] = dfAwardDef.loc[dfAwardDef["ColumnName"] == this_col_name, "Label" + str(award_num + 1)]
    add_table_dataframe(ojsfile, "AwardListDropdowns", "AwardListDropdowns", j_awards_df)

    # Robot Game
    rg_awards = int(s1["P_AWD_RG"])
    print(f"There are {rg_awards} robot game awards")
    rg_awards_df = pd.DataFrame(columns=["Robot Game Awards"])
    for rg_award_num in range(0, rg_awards):
        rg_awards_df.loc[len(rg_awards_df)] = [config["TOURNAMENT_AWARDS"]["Robot Game"]["dropdown"]["p" + str(rg_award_num + 1)]]
    add_table_dataframe(ojsfile, "AwardListDropdowns", "RobotGameAwards", rg_awards_df)


def set_up_award_worksheet_div(tournament: pd.Series):
    for d in ["D1", "D2"]:
        ojsfile = (
            dir_path
            + "\\tournaments\\"
            + tournament["Short Name"]
            + "\\"
            + tournament[d + "_OJS"]
        )
        print(Fore.RESET)
        print(f"Setting up Tournament Division Awards for {tournament["Short Name"]} {d}")
        divawards: pd.DataFrame = dfDivAwards[
            (dfDivAwards["Tournament"] == tournament["Short Name"])
            & (dfDivAwards["Div"] == d)
        ]
        print(divawards)

        # Robot Game
        if tournament[d + "_OJS"] is None:
            print(Fore.RESET)
            print(f"Nothing for tournament[{d}]")
            continue
        
        rg_awards = divawards["P_AWD_RG"].iloc[0]
        print(f"There are {rg_awards} robot game awards")
        rg_awards_df = pd.DataFrame(columns=["Robot Game Awards"])
        for rg_award_num in range(0, rg_awards):
            rg_awards_df.loc[len(rg_awards_df)] = dfAwardDef.loc[dfAwardDef["ColumnName"] == "P_AWD_RG", "Label" + str(rg_award_num + 1)]
            # rg_awards_df.loc[len(rg_awards_df)] = [config["TOURNAMENT_AWARDS"]["Robot Game"]["dropdown"]["p" + str(rg_award_num + 1)]]
        add_table_dataframe(ojsfile, "AwardListDropdowns", "RobotGameAwards", rg_awards_df)

        # Other judged awards
        j_cols = divawards.filter(regex=f"^J_")
        # print(str(config).replace("'", '"'))
        j_awards_df = pd.DataFrame(columns=["Award"])
        for this_col_name, series in j_cols.items():
            print(f"Working the {this_col_name} awards")
            for award_num in range(0, series.iloc[0]):
                # print("Adding new judged award")
                # print(find_item_by_col_name(d=config, col_name=this_col_name))
                j_awards_df.loc[len(j_awards_df)] = dfAwardDef.loc[dfAwardDef["ColumnName"] == this_col_name, "Label" + str(award_num + 1)]
        add_table_dataframe(ojsfile, "AwardListDropdowns", "AwardListDropdowns", j_awards_df)


def do_division_awards(div: str, d: Dict, tournament: pd.Series):
    d["FILES"]["ojs1"] = tournament["D1_OJS"]
    d[div] = {}
    d[div]["notes"] = "If your tournament uses divisions, set up the awards here. If you set div1 to false above, this section will be ignored. Same for div2 and the D2 section below."
    print(f"Writing division awards for {div} to the config json file")
    print(dfDivAwards.loc[(dfDivAwards["Tournament"] == tournament["Short Name"]) & (dfDivAwards["Div"] == div)])
    awd_cols = dfDivAwards.filter(regex=r'^(?:J_|P_)').loc[(dfDivAwards["Tournament"] == tournament["Short Name"]) & (dfDivAwards["Div"] == div)]
    for this_awd_col, this_awd_qty in awd_cols.items():
        # build the award entry using qty and dropdown labels
        this_awd_name = dfAwardDef.loc[dfAwardDef["ColumnName"] == this_awd_col, "Name"]
        awd_conf = config["TOURNAMENT_AWARDS"].get(this_awd_name, {})

        # normalize quantity to an int (safe fallback to 0)
        try:
            qty = int(this_awd_qty.iloc[0])
        except Exception:
            qty = 0

        # store qty as string to match desired JSON output and build places as list of dicts
        d[div][this_awd_name] = {"qty": str(qty), "places": []}

        dropdown = awd_conf.get("dropdown", {}) if isinstance(awd_conf, dict) else {}
        places = []
        for i in range(qty):
            pkey = f"p{i+1}"
            label = dropdown.get(pkey, "")  # fallback to empty string if missing
            places.append({"place": str(i + 1), "label": label})

        d[div][this_awd_name]["places"] = places

def set_up_meta_worksheet(tournament: pd.Series):
    print(Fore.RESET)
    print(f"Setting up meta worksheet for {tournament["Short Name"]}")
    print(tournament)
    scriptfile = (
        dir_path
        + "\\tournaments\\"
        + tournament["Short Name"]
        + "\\closing_ceremony.html"
    )
    dfMeta = pd.DataFrame(columns=["Key", "Value"])
    dfMeta.loc[len(dfMeta)] = {"Key": "Tournament Year", "Value": config["season_yr"]}
    dfMeta.loc[len(dfMeta)] = {"Key": "FLL Season Title", "Value": config["season_name"]}
    dfMeta.loc[len(dfMeta)] = {"Key": "Tournament Short Name", "Value": tournament["Short Name"]}
    dfMeta.loc[len(dfMeta)] = {"Key": "Tournament Long Name", "Value": tournament["Long Name"]}
    dfMeta.loc[len(dfMeta)] = {"Key": "Completed Script File", "Value": scriptfile}
    dfMeta.loc[len(dfMeta)] = {"Key": "Using Divisions", "Value": using_divisions}

    ojsfile = (
        dir_path
        + "\\tournaments\\"
        + tournament["Short Name"]
        + "\\"
        + tournament["D1_OJS"]
    )
    print(Fore.RESET)
    add_table_dataframe(xlsx_path=ojsfile,sheet_name="Meta", table_name="Meta", data=dfMeta)


def set_up_meta_worksheet_div(tournament: pd.Series):
    for d in ["D1", "D2"]:
        print(Fore.RESET)
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
        dfMeta.loc[len(dfMeta)] = {"Key": "FLL Season Title", "Value": config["season_name"]}
        dfMeta.loc[len(dfMeta)] = {"Key": "Tournament Short Name", "Value": tournament["Short Name"]}
        dfMeta.loc[len(dfMeta)] = {"Key": "Tournament Long Name", "Value": tournament["Long Name"]}
        dfMeta.loc[len(dfMeta)] = {"Key": "Completed Script File", "Value": scriptfile}
        dfMeta.loc[len(dfMeta)] = {"Key": "Using Divisions", "Value": using_divisions}
        dfMeta.loc[len(dfMeta)] = {"Key": "Division", "Value": d}
        dfMeta.loc[len(dfMeta)] = {"Key": "Advancing", "Value": tournament["ADV"]}

        if tournament[d + "_OJS"] is not None:
            ojsfile = (
                dir_path
                + "\\tournaments\\"
                + tournament["Short Name"]
                + "\\"
                + tournament[d + "_OJS"]
            )
            print(Fore.RESET)
            add_table_dataframe(xlsx_path=ojsfile,sheet_name="Meta", table_name="Meta", data=dfMeta)


def copy_team_numbers(
    source_sheet: Worksheet, target_sheet: Worksheet, target_start_row: int
):
    source_start_row = 3

    column = "A"
    last_row = 0
    print(Fore.RESET)
    print(f"Copying team numbers to {target_sheet}")
    # Iterate through the rows in the specified column
    for row in range(1, source_sheet.max_row + 1):
        if source_sheet[f"{column}{row}"].value is not None:
            last_row = row
    team_count = last_row - source_start_row + 1
    col = 1  # Team number is always in column 1 ('A')
    # itterate over the source rows. The dest row may not always align with
    # the source row. Some sheets start on row 3, some start on 2
    current_target_row = target_start_row
    for row in range(source_start_row, source_start_row + team_count + 1):
        cell_value = source_sheet.cell(row=row, column=col).value
        target_sheet.cell(row=current_target_row, column=col).value = cell_value
        current_target_row += 1


def protect_worksheets(tournament: pd.Series):
    for d in ["D1", "D2"]:
        if tournament[d + "_OJS"] is None:
            print(Fore.RESET)
            print(f'*-*-*-* No division {d} to check for {tournament["Short Name"]}')
            continue
        print(Fore.RESET)
        print(f'Protecting {d} {tournament["Short Name"]}')
        ojsfile = (
            dir_path
            + "\\tournaments\\"
            + tournament["Short Name"]
            + "\\"
            + tournament[d + "_OJS"]
        )
        ojs_book = load_workbook(ojsfile, read_only=False, keep_vba=True)
        for ws in ojs_book.worksheets:
            # print(Fore.RESET)
            # print(f"Protecting {ws}")
            ws.protection.selectLockedCells = True
            ws.protection.selectUnlockedCells = False
            ws.protection.formatCells = False
            ws.protection.formatColumns = False
            ws.protection.formatRows = False
            ws.protection.autoFilter = False
            ws.protection.sort = False
            ws.protection.set_password(SHEET_PASSWORD)
            # ws.protection.sheet = True
            ws.protection.enable()
            print(f"{ws} is protected")

        ojs_book.save(ojsfile)


def resize_worksheets(tournament: pd.Series):
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
    for d in ["D1", "D2"]:
        sheet_tables = zip(worksheetNames, worksheetTables, worksheet_start_row)
        if tournament[d + "_OJS"] is None:
            print(Fore.RESET)
            print(f'*-*-*-* No division {d} to check for {tournament["Short Name"]}')
            continue

        print(Fore.RESET)
        print(f'Resizing {d} {tournament["Short Name"]}')
        ojsfile = (
            dir_path
            + "\\tournaments\\"
            + tournament["Short Name"]
            + "\\"
            + tournament[d + "_OJS"]
        )
        ojs_book = load_workbook(ojsfile, read_only=False, keep_vba=True)
        divassignees: pd.DataFrame = dfAssignments[
            (dfAssignments["Tournament"] == tournament["Short Name"])
            & (dfAssignments["Div"] == d)
        ]

        # copy the team number data to each of the worksheets
        for s, t, r in sheet_tables:
            ws = ojs_book[s]
            tapi_sheet = ojs_book["Team and Program Information"]
            # first, copy the team numbers over
            copy_team_numbers(
                source_sheet=tapi_sheet, target_sheet=ws, target_start_row=r
            )

        # Resize the tables
        sheet_tables = zip(worksheetNames, worksheetTables, worksheet_start_row)
        for s, t, r in sheet_tables:
            print(Fore.RESET)
            print(f"{s}, {r}, {t}")
            ws = ojs_book[s]
            table: Table = ws.tables[t]
            table_range: str = table.ref
            start_cell = table_range.split(":")[0]
            start_row = int(re.findall(r"\d", start_cell)[0])
            colonPos = (table.ref).find(":")
            # print(f'Resizing the {d} {t} table {table.ref}, {table.ref[:colonPos]}, {re.sub(r'\d', '', table.ref[colonPos + 1])}, {str(len(divassignees) + 2)}.')
            table.ref = (
                table.ref[:colonPos]
                + ":"
                + re.sub(r"\d", "", table.ref[colonPos + 1])
                + str(start_row + len(divassignees))
            )
            # print(f'New table.ref = {table.ref}')
            ws.delete_rows(idx=start_row + len(divassignees) + 1, amount=200)
            # ws.protection.sheet = True

        ojs_book.save(ojsfile)

def copy_award_def(tournament: pd.Series):
    if using_divisions:
        for d in ["D1", "D2"]:
            ojsfile = (
                dir_path
                + "\\tournaments\\"
                + tournament["Short Name"]
                + "\\"
                + tournament[d + "_OJS"]
            )
            add_table_dataframe(ojsfile, "AwardDef", "AwardDef", dfAwardDef)
    else:
        ojsfile = (
            dir_path
            + "\\tournaments\\"
            + tournament["Short Name"]
            + "\\"
            + tournament["D1_OJS"]
        )
        add_table_dataframe(ojsfile, "AwardDef", "AwardDef", dfAwardDef)


def add_conditional_formats(tournament: pd.Series):
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
    for d in ["D1", "D2"]:
        if tournament[d + "_OJS"] is None:
            continue
        print(Fore.RESET)
        print(f'Adding conditional formats to {d} {tournament["Short Name"]}')
        ojsfile = (
            dir_path
            + "\\tournaments\\"
            + tournament["Short Name"]
            + "\\"
            + tournament[d + "_OJS"]
        )
        ojs_book = load_workbook(ojsfile, read_only=False, keep_vba=True)
        ws = ojs_book["Results and Rankings"]
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
        ojs_book.save(ojsfile)


def hide_worksheets(tournament: pd.Series):
    worksheetNames = ["Data Validation", "Meta", "AwardListDropdowns", "AwardDef"]
    for d in ["D1", "D2"]:
        if tournament[d + "_OJS"] is None:
            continue
        print(f'Hiding worksheets in {d} {tournament["Short Name"]}')
        ojsfile = (
            dir_path
            + "\\tournaments\\"
            + tournament["Short Name"]
            + "\\"
            + tournament[d + "_OJS"]
        )
        ojs_book = load_workbook(ojsfile, read_only=False, keep_vba=True)

        for sheetname in worksheetNames:
            ws = ojs_book[sheetname]
            ws.sheet_state = "hidden"

        ojs_book.save(ojsfile)


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

tournament_file: str = dir_path + "\\" + config['filename']
template_file: str = dir_path + "\\" + config['tournament_template']
extrafilelist: list[str] = config['copy_file_list']

# Make sure the extra files exist
for filename in extrafilelist:
    try:
        if os.path.exists(dir_path + "\\" + filename):
            print(Fore.GREEN +
                f"{filename}... CHECK!",
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
    print_error(f"Could not open the tournament file: {tournament_file}. Check to make sure it is not open in Excel.", e)

try:
    dictSeasonInfo: pd.DataFrame = read_table_as_dict(tournament_file, "SeasonInfo", "SeasonInfo")
    using_divisions: bool = dictSeasonInfo["Divisions"]
except Exception as e:
    print_error(f"Could not read the SeasonInfo table.", e)

try:
    dfTournaments = read_table_as_df(tournament_file, "Tournaments", "TournamentList")
except Exception as e:
    print_error(f"Could not read the tournament worksheet.", e)

try:
    dfDivAwards = read_table_as_df(tournament_file, "DivAwards", "AwardListDiv")
except Exception as e:
    print_error(f"Could not read the divAwards worksheet.", e)

try:
    dfAwardDef = read_table_as_df(tournament_file, "AwardDef", "AwardDef")
except Exception as e:
    print_error(f"Could not read the AwardDef worksheet.", e)

try:
    dfAssignments = read_table_as_df(tournament_file, "Assignments", "Assignments")
    tourn_array: list[str] = []
    for index, row in dfTournaments.iterrows():
        tourn_array.append(row["Short Name"])
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

# Now that we have all of the info for the tournaments, loop through and
# start building the OJS files and folders
if using_divisions:
    for index, row in dfDivAwards.iterrows():
        pass

for index, row in dfTournaments.iterrows():
    newpath = dir_path + "\\tournaments\\" + row["Short Name"]
    create_folder(newpath)
    copy_files(row)
    set_up_tapi_worksheet(row.fillna(0))
    if using_divisions:
        set_up_award_worksheet_div(row.fillna(0))
        set_up_meta_worksheet_div(row.fillna(0)) #TODO
    else:
        set_up_meta_worksheet(row.fillna(0)) #TODO
    set_up_award_worksheet(row.fillna(0))
    add_conditional_formats(row)
    copy_award_def(row.fillna(0))
    hide_worksheets(row)
    resize_worksheets(row)
    protect_worksheets(row)


print(Fore.GREEN)
input(f"All done. Created OJS workbooks for {len(dfTournaments)} tournament(s). Press enter to quit...")
sys.exit(1)
