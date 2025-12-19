"""Worksheet setup functions for tournament OJS files."""

import os
import logging
import pandas as pd
import numpy as np
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.cell_range import CellRange
from copy import copy as _copy

from .constants import *
from .excel_operations import add_table_dataframe, _to_int
from .logger import print_error


logger = logging.getLogger("ojs_builder")


def set_up_tapi_worksheet(
    tournament: pd.Series,
    book: Workbook,
    dfAssignments: pd.DataFrame,
    using_divisions: bool
) -> None:
    """Populate the 'Team and Program Information' table.

    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        dfAssignments: DataFrame containing team assignments
        using_divisions: Boolean indicating if divisions are used
    """
    d = ""
    logger.info(f"Setting up Team and Program Information for {tournament[COL_SHORT_NAME]}")
    
    if using_divisions:
        d = tournament[COL_DIVISION]
        if isinstance(tournament[COL_OJS_FILENAME], float):
            print_error(logger, f"Invalid OJS filename for tournament {tournament[COL_SHORT_NAME]}: "
                       "expected string, got float (possibly missing value)")
        
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments[COL_SHORT_NAME] == tournament[COL_SHORT_NAME])
            & (dfAssignments[COL_DIVISION] == d)
        ]
        logger.info(f"Found {len(assignees)} teams in {tournament[COL_SHORT_NAME]} {d}")
    else:
        if isinstance(tournament[COL_OJS_FILENAME], float):
            print_error(logger, f"Invalid OJS filename for tournament {tournament[COL_SHORT_NAME]}: "
                       "expected string, got float (possibly missing value)")
        
        assignees: pd.DataFrame = dfAssignments[
            (dfAssignments[COL_SHORT_NAME] == tournament[COL_SHORT_NAME])
        ]
        logger.info(f"Found {len(assignees)} teams in {tournament[COL_SHORT_NAME]}")

    # Validate required columns
    keep = [COL_TEAM_NUMBER, COL_TEAM_NAME, COL_COACH_NAME]
    keep_safe = [c for c in keep if c in assignees.columns]
    
    if len(keep_safe) != len(keep):
        missing = set(keep) - set(keep_safe)
        print_error(logger, f"Missing required columns in assignments: {', '.join(missing)}")
    
    assignees = assignees[keep_safe]
    assignees[COL_POD_NUMBER] = 0
    sorted_assignees = assignees.sort_values(by=COL_TEAM_NUMBER, ascending=True)
    
    add_table_dataframe(book, SHEET_TEAM_INFO, TABLE_TEAM_LIST, sorted_assignees)


def set_up_award_worksheet(
    tournament: pd.Series,
    book: Workbook,
    dfAwardDef: pd.DataFrame,
    using_divisions: bool
) -> None:
    """Prepare award tables used by OJS dropdowns and closing scripts.

    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        dfAwardDef: DataFrame containing award definitions
        using_divisions: Boolean indicating if divisions are used
    """
    thisDiv = tournament[COL_DIVISION] if using_divisions else ""
    logger.info(f"Setting up awards for {tournament[COL_SHORT_NAME]} {thisDiv}")

    # Robot Game awards
    rg_raw = (
        tournament.get(AWARD_COLUMN_ROBOT_GAME)
        if hasattr(tournament, "get")
        else tournament[AWARD_COLUMN_ROBOT_GAME]
    )
    try:
        rg_awards = int(rg_raw) if not pd.isna(rg_raw) else 0
    except (ValueError, TypeError) as e:
        print_error(logger, f"Invalid robot game award count for {tournament[COL_SHORT_NAME]}: {rg_raw}", e)
        
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
    logger.debug(f"Added {len(rg_awards_df)} robot game awards and {len(j_awards_df)} judged awards")


def set_up_meta_worksheet(
    tournament: pd.Series,
    book: Workbook,
    config: dict,
    dir_path: str,
    using_divisions: bool
) -> None:
    """Populate the metadata worksheet with tournament information.
    
    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        config: Dictionary containing season configuration
        dir_path: Base directory path
        using_divisions: Boolean indicating if divisions are used
    """
    d = tournament[COL_DIVISION] if using_divisions else ""
    logger.info(f"Setting up metadata for {tournament[COL_SHORT_NAME]} {d}")
    
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
            print_error(logger, f"Missing required configuration key: '{key}' in {CONFIG_FILENAME}")
    
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
    
    if tournament[COL_OJS_FILENAME] is not None:
        add_table_dataframe(book, SHEET_META, TABLE_META, dfMeta, debug=False)


def copy_team_numbers(
    source_sheet: Worksheet,
    target_sheet: Worksheet,
    target_start_row: int,
    source_start_row: int = 3,
    debug: bool = False,
) -> int:
    """Copy team numbers from column A of source to target sheet.
    
    Args:
        source_sheet: Source worksheet
        target_sheet: Target worksheet
        target_start_row: Row to start writing in target
        source_start_row: Row to start reading from source
        debug: If True, log diagnostic information
        
    Returns:
        Number of rows copied
    """
    col = 1
    max_row = source_sheet.max_row

    if debug:
        logger.debug(
            f"copy_team_numbers: source={source_sheet.title}, target={target_sheet.title}, "
            f"max_row={max_row}, source_start={source_start_row}, target_start={target_start_row}"
        )

    last_row = None
    for r in range(source_start_row, max_row + 1):
        v = source_sheet.cell(row=r, column=col).value
        if v is not None and v != "":
            last_row = r

    if last_row is None:
        logger.debug("copy_team_numbers: no non-empty rows found")
        return 0

    dest_row = target_start_row
    copied = 0
    for r in range(source_start_row, last_row + 1):
        cell_value = source_sheet.cell(row=r, column=col).value
        
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

    logger.debug(f"copy_team_numbers: copied {copied} rows")
    return copied


# Copy formatting helpers and other worksheet functions here...
# (I'll include the critical ones - let me know if you want all of them)

def protect_worksheets(tournament: pd.Series, book: Workbook) -> None:
    """Apply protection settings to every worksheet.
    
    Args:
        tournament: Tournament row
        book: Workbook to protect
    """
    logger.info(f"Protecting worksheets for {tournament[COL_OJS_FILENAME]}")
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
    logger.debug("Worksheet protection applied")


def hide_worksheets(tournament: pd.Series, book: Workbook) -> None:
    """Hide utility worksheets.
    
    Args:
        tournament: Tournament row
        book: Workbook
    """
    logger.info(f"Hiding worksheets for {tournament[COL_OJS_FILENAME]}")
    worksheetNames = ["Data Validation", SHEET_META, SHEET_AWARD_DROPDOWNS, SHEET_AWARD_DEF]
    
    for sheetname in worksheetNames:
        if sheetname in book.sheetnames:
            ws = book[sheetname]
            ws.sheet_state = "hidden"
            logger.debug(f"Hid worksheet: {sheetname}")
