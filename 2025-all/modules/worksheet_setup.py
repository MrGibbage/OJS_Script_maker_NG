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

from .constants import (
    SHEET_PASSWORD, REQUIRED_COLUMNS,
    COL_TEAM_NUMBER, COL_TEAM_NAME, COL_COACH_NAME, COL_POD_NUMBER,
    COL_SHORT_NAME, COL_LONG_NAME, COL_OJS_FILENAME, COL_DIVISION, COL_ADVANCING,
    SHEET_TEAM_INFO, SHEET_AWARD_DROPDOWNS, SHEET_META, SHEET_AWARD_DEF,
    SHEET_ROBOT_GAME, SHEET_INNOVATION, SHEET_ROBOT_DESIGN, SHEET_CORE_VALUES, SHEET_RESULTS,
    TABLE_TEAM_LIST, TABLE_ROBOT_GAME_AWARDS, TABLE_AWARD_DROPDOWNS, TABLE_META, TABLE_AWARD_DEF,
    TABLE_ROBOT_GAME, TABLE_INNOVATION, TABLE_ROBOT_DESIGN, TABLE_CORE_VALUES, TABLE_TOURNAMENT_DATA,
    FILE_CLOSING_CEREMONY, AWARD_COLUMN_PREFIX_JUDGED, AWARD_COLUMN_ROBOT_GAME, AWARD_LABEL_PREFIX
)
from .excel_operations import add_table_dataframe, _to_int
from .logger import print_error


logger = logging.getLogger("ojs_builder")


def set_up_tapi_worksheet(
    tournament: pd.Series,
    book: Workbook,
    dfAssignments: pd.DataFrame,
    using_divisions: bool
) -> bool:
    """Populate the 'Team and Program Information' table.

    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        dfAssignments: DataFrame containing team assignments
        using_divisions: Boolean indicating if divisions are used
        
    Returns:
        True if successful, False if no teams assigned (should skip this tournament)
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

    # Check if there are any teams assigned
    if len(assignees) == 0:
        logger.warning(f"No teams assigned to {tournament[COL_SHORT_NAME]} {d} - skipping")
        return False

    # Validate required columns exist
    keep = [COL_TEAM_NUMBER, COL_TEAM_NAME, COL_COACH_NAME]
    keep_safe = [c for c in keep if c in assignees.columns]
    
    if len(keep_safe) != len(keep):
        missing = set(keep) - set(keep_safe)
        print_error(logger, f"Missing required columns in assignments: {', '.join(missing)}")
    
    assignees = assignees[keep_safe]
    assignees[COL_POD_NUMBER] = 0
    sorted_assignees = assignees.sort_values(by=COL_TEAM_NUMBER, ascending=True)
    
    add_table_dataframe(book, SHEET_TEAM_INFO, TABLE_TEAM_LIST, sorted_assignees)
    return True


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
    
    # Get available Label columns dynamically
    label_columns = [col for col in dfAwardDef.columns if col.startswith(AWARD_LABEL_PREFIX)]
    logger.debug(f"Found {len(label_columns)} label columns: {label_columns}")

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
        
        # Check if this Label column exists
        if thisLabel not in label_columns:
            logger.warning(f"Label column '{thisLabel}' not found in AwardDef table, using blank value")
            thisValue = None
        else:
            sel = dfAwardDef.loc[dfAwardDef["ColumnName"] == AWARD_COLUMN_ROBOT_GAME, thisLabel]
            try:
                thisValue = sel.iat[0]
            except (IndexError, KeyError):
                logger.warning(f"Could not find value for {AWARD_COLUMN_ROBOT_GAME} in column {thisLabel}")
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
            
            # Check if this Label column exists
            if label_col not in label_columns:
                logger.warning(f"Label column '{label_col}' not found in AwardDef table, using blank value")
                thisValue = None
            else:
                sel = dfAwardDef.loc[dfAwardDef["ColumnName"] == this_col_name, label_col]
                try:
                    thisValue = sel.iat[0]
                except (IndexError, KeyError):
                    logger.warning(f"Could not find value for {this_col_name} in column {label_col}")
                    thisValue = None
            
            j_awards_df.loc[len(j_awards_df)] = [thisValue]
            
    add_table_dataframe(book, SHEET_AWARD_DROPDOWNS, TABLE_AWARD_DROPDOWNS, j_awards_df)
    logger.debug(f"Added {len(rg_awards_df)} robot game awards and {len(j_awards_df)} judged awards")


def set_up_meta_worksheet(
    tournament: pd.Series,
    book: Workbook,
    config: dict,
    tournament_folder: str,
    using_divisions: bool
) -> None:
    """Populate the metadata worksheet with tournament information.
    
    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        config: Dictionary containing season configuration
        tournament_folder: Root folder where tournament subfolders are created
        using_divisions: Boolean indicating if divisions are used
    """
    d = tournament[COL_DIVISION] if using_divisions else ""
    logger.info(f"Setting up metadata for {tournament[COL_SHORT_NAME]} {d}")
    
    scriptfile = os.path.join(
        tournament_folder,
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


def _copy_data_validations_for_range(
    ws: Worksheet,
    start_col_idx: int,
    end_col_idx: int,
    first_data_row: int,
    new_end_row: int,
):
    """Duplicate data validations from template row to new range.
    
    Args:
        ws: Worksheet to modify
        start_col_idx: Starting column index
        end_col_idx: Ending column index
        first_data_row: Template data row number
        new_end_row: New ending row number
    """
    try:
        existing = list(ws.data_validations.dataValidation)
    except Exception:
        existing = []

    for dv in existing:
        try:
            for rng in dv.ranges:
                try:
                    cr = CellRange(str(rng))
                except Exception:
                    continue

                if not (cr.min_row <= first_data_row <= cr.max_row):
                    continue

                orig_min_col = cr.min_col
                orig_max_col = cr.max_col
                new_min_col = max(orig_min_col, start_col_idx)
                new_max_col = min(orig_max_col, end_col_idx)

                if new_min_col > new_max_col:
                    continue

                new_range = f"{get_column_letter(new_min_col)}{first_data_row}:{get_column_letter(new_max_col)}{new_end_row}"

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

                try:
                    newdv.add(new_range)
                    ws.add_data_validation(newdv)
                except Exception:
                    continue
        except Exception:
            continue


def _extend_conditional_formatting_for_range(
    ws: Worksheet,
    start_col_letter: str,
    end_col_letter: str,
    first_data_row: int,
    new_end_row: int,
):
    """Duplicate conditional formatting rules to new range.
    
    Args:
        ws: Worksheet to modify
        start_col_letter: Starting column letter
        end_col_letter: Ending column letter
        first_data_row: Template data row number
        new_end_row: New ending row number
    """
    try:
        cf_rules = getattr(ws.conditional_formatting, "_cf_rules", {})
    except Exception:
        cf_rules = {}

    new_range = f"{start_col_letter}{first_data_row}:{end_col_letter}{new_end_row}"

    for key, rules in list(cf_rules.items()):
        try:
            for sub in str(key).split():
                try:
                    cr = CellRange(sub)
                except Exception:
                    continue
                if cr.min_row <= first_data_row <= cr.max_row:
                    for rule in rules:
                        try:
                            ws.conditional_formatting.add(new_range, rule)
                        except Exception:
                            continue
        except Exception:
            continue


def resize_worksheets(
    tournament: pd.Series,
    book: Workbook,
    dfAssignments: pd.DataFrame,
    using_divisions: bool
) -> None:
    """Resize the main result/input tables to match team count.

    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        dfAssignments: DataFrame containing team assignments
        using_divisions: Boolean indicating if divisions are used
    """
    logger.info(f"Resizing worksheets for {tournament[COL_OJS_FILENAME]}")
    
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

    # Copy team numbers to each worksheet
    sheet_tables = zip(worksheetNames, worksheetTables, worksheet_start_row)
    copied_counts: dict[str, int] = {}
    for s, t, r in sheet_tables:
        if s in book.sheetnames:
            ws = book[s]
            tapi_sheet = book[SHEET_TEAM_INFO]
            copied = copy_team_numbers(
                source_sheet=tapi_sheet,
                target_sheet=ws,
                target_start_row=r,
                debug=False,
            )
            copied_counts[s] = copied

    # Resize the tables
    sheet_tables = zip(worksheetNames, worksheetTables, worksheet_start_row)
    for s, t, r in sheet_tables:
        if s not in book.sheetnames:
            continue
        ws = book[s]
        table: Table = ws.tables[t]
        table_range: str = table.ref

        start_cell, end_cell = table_range.split(":")
        start_col_letter, start_row_num = coordinate_from_string(start_cell)
        end_col_letter, end_row_num = coordinate_from_string(end_cell)

        rows_for_table = copied_counts.get(s, len(assignees))
        new_end_row = start_row_num + rows_for_table

        table.ref = f"{start_col_letter}{start_row_num}:{end_col_letter}{new_end_row}"
        logger.debug(f"Resized table {t}: {table.ref}")

        # Copy formulas
        first_data_row = start_row_num + 1
        start_col_idx = column_index_from_string(start_col_letter)
        end_col_idx = column_index_from_string(end_col_letter)
        
        for col_idx in range(start_col_idx + 1, end_col_idx + 1):
            template = ws.cell(row=first_data_row, column=col_idx).value
            if isinstance(template, str) and template.startswith("="):
                for rr in range(first_data_row, new_end_row + 1):
                    ws.cell(row=rr, column=col_idx).value = template

        # Copy cell protection
        for col_idx in range(start_col_idx, end_col_idx + 1):
            template_cell = ws.cell(row=first_data_row, column=col_idx)
            tpl_prot = getattr(template_cell, "protection", None)
            if tpl_prot is None:
                continue
            for rr in range(first_data_row, new_end_row + 1):
                tgt = ws.cell(row=rr, column=col_idx)
                try:
                    tgt.protection = _copy(tpl_prot)
                except Exception:
                    pass

        # Copy cell styles
        for col_idx in range(start_col_idx, end_col_idx + 1):
            template_cell = ws.cell(row=first_data_row, column=col_idx)
            for rr in range(first_data_row, new_end_row + 1):
                tgt = ws.cell(row=rr, column=col_idx)
                try:
                    if hasattr(template_cell, "_style"):
                        tgt._style = _copy(template_cell._style)
                    tgt.number_format = template_cell.number_format
                    tgt.font = _copy(template_cell.font)
                    tgt.fill = _copy(template_cell.fill)
                    tgt.alignment = _copy(template_cell.alignment)
                    tgt.border = _copy(template_cell.border)
                except Exception:
                    pass

        # Copy row height
        try:
            template_height = ws.row_dimensions[first_data_row].height
            if template_height is not None:
                for rr in range(first_data_row, new_end_row + 1):
                    ws.row_dimensions[rr].height = template_height
        except Exception:
            pass

        # Copy data validations and conditional formatting
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

        # Delete extra rows
        ws.delete_rows(idx=new_end_row + 1, amount=200)

    logger.debug("Worksheet resizing complete")


def copy_award_def(tournament: pd.Series, book: Workbook, dfAwardDef: pd.DataFrame) -> None:
    """Add the award definitions table to the workbook.
    
    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
        dfAwardDef: DataFrame containing award definitions
    """
    logger.debug("Adding award definitions table")
    add_table_dataframe(book, SHEET_AWARD_DEF, TABLE_AWARD_DEF, dfAwardDef)


def add_conditional_formats(tournament: pd.Series, book: Workbook) -> None:
    """Add conditional formatting rules to the Results and Rankings worksheet.
    
    Args:
        tournament: A pandas Series representing the tournament row
        book: An open openpyxl Workbook object
    """
    logger.info(f"Adding conditional formats for {tournament[COL_OJS_FILENAME]}")
    
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
    
    ws = book[SHEET_RESULTS]
    
    # Award completion
    ws.conditional_formatting.add(
        "O2",
        FormulaRule(
            formula=["COUNTA(AwardList!$A$2:$A$35)=COUNTA($O$3:$O$288)"],
            stopIfTrue=False,
            fill=greenAwardFill,
        ),
    )
    
    # Advancing count
    ws.conditional_formatting.add(
        "P2",
        FormulaRule(
            formula=['COUNTIF($P:$P,"Yes")=Meta!$B$13'],
            stopIfTrue=False,
            fill=greenAdvFill,
        ),
    )
    
    # Robot Game medals
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
    
    logger.debug("Conditional formatting applied")


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
