"""File and folder operations for tournament setup."""

import os
import json
import shutil
import logging
from typing import Any
import pandas as pd
from openpyxl import load_workbook

from .constants import (
    COL_SHORT_NAME, 
    COL_LONG_NAME,
    COL_OJS_FILENAME, 
    COL_DATE, 
    COL_DIVISION,
    COL_COLUMN_NAME, 
    COL_DIV_AWARD,
    COL_SCRIPT_TAG_D1,
    COL_SCRIPT_TAG_D2,
    COL_SCRIPT_TAG_NODIV,
    AWARD_COLUMN_PREFIX_JUDGED,
    AWARD_COLUMN_ROBOT_GAME,
    AWARD_LABEL_PREFIX,
    SHEET_TEAM_INFO
)
from .logger import print_error


logger = logging.getLogger("ojs_builder")


def _remove_note_keys(obj: Any) -> Any:
    """Recursively strip out any dict keys beginning with 'note' (case-insensitive).
    
    This allows JSON config files to include 'note' keys for human-readable comments.
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
    """Load JSON from path and return a copy with 'note' keys removed.
    
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
        logger.debug(f"Loading configuration from: {path}")
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        logger.info(f"✓ Successfully loaded configuration")
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError(
            f"Invalid JSON syntax in configuration file: {path}",
            e.doc,
            e.pos
        ) from e
    
    return _remove_note_keys(data)


def create_folder(newpath: str) -> None:
    """Create directory if it does not exist.
    
    Args:
        newpath: Path to the directory to create
        
    Raises:
        Exits via print_error if directory creation fails
    """
    if not newpath or not newpath.strip():
        print_error(logger, "Cannot create folder: path is empty or None")
        
    if os.path.exists(newpath):
        logger.debug(f"Folder already exists: {newpath}")
        return
    
    try:
        os.makedirs(newpath)
        logger.info(f"✓ Created folder: {newpath}")
    except PermissionError as e:
        print_error(
            logger, 
            f"Permission denied when creating directory: {newpath}", 
            e,
            error_type='permission_denied',
            context={'filename': newpath}
        )
    except OSError as e:
        print_error(logger, f"OS error when creating directory: {newpath}", e)
    except Exception as e:
        print_error(logger, f"Unexpected error creating directory: {newpath}", e)


def copy_files(
    item: pd.Series,
    dir_path: str,
    template_file: str,
    common_files: list[dict],
    divisions_only_files: list[dict],
    no_divisions_only_files: list[dict],
    tournament_folder: str,
    using_divisions: bool = False
) -> None:
    """Copy extra files and OJS template into tournament folder.
    
    Handles three types of file lists:
    - common_files: Always copied (e.g., executables, PDFs)
    - divisions_only_files: Only copied when using_divisions=True
    - no_divisions_only_files: Only copied when using_divisions=False
    
    Each file list contains dicts with {"source": "...", "dest": "..."}.
    This allows source files to have distinct names (e.g., script_template-with-divisions.html.jinja)
    while destination files have consistent names (e.g., script_template.html.jinja).
    
    Args:
        item: Tournament row with 'Short Name' and 'OJS_FileName'
        dir_path: Base directory containing source files
        template_file: Path to OJS template workbook
        common_files: List of {source, dest} dicts for files always copied
        divisions_only_files: List of {source, dest} dicts for division tournaments only
        no_divisions_only_files: List of {source, dest} dicts for non-division tournaments only
        tournament_folder: Root folder where tournament subfolders are created
        using_divisions: Whether the tournament uses divisions
        
    Raises:
        Exits via print_error if any copy operation fails
    """
    if COL_SHORT_NAME not in item or not item[COL_SHORT_NAME]:
        print_error(logger, f"Tournament row missing required field '{COL_SHORT_NAME}'")
        
    if COL_OJS_FILENAME not in item or not item[COL_OJS_FILENAME]:
        print_error(logger, f"Tournament row missing required field '{COL_OJS_FILENAME}' for {item[COL_SHORT_NAME]}")
    
    logger.info(f"Copying files for tournament: {item[COL_SHORT_NAME]}")
    dest_folder = os.path.join(tournament_folder, item[COL_SHORT_NAME])
    
    # Determine which file lists to process
    files_to_copy = common_files.copy()
    if using_divisions:
        files_to_copy.extend(divisions_only_files)
        logger.debug(f"Using divisions mode: copying {len(divisions_only_files)} division-specific files")
    else:
        files_to_copy.extend(no_divisions_only_files)
        logger.debug(f"Using non-divisions mode: copying {len(no_divisions_only_files)} non-division-specific files")
    
    # Copy all files
    for file_mapping in files_to_copy:
        try:
            source_file = file_mapping["source"]
            dest_file = file_mapping["dest"]
            
            source_path = os.path.join(dir_path, source_file)
            if not os.path.exists(source_path):
                print_error(
                    logger,
                    f"Source file not found: {source_path}",
                    error_type='missing_file',
                    context={'filename': source_file, 'directory': dir_path}
                )
            
            dest_path = os.path.join(dest_folder, dest_file)
            
            # Log if overwriting existing file
            if os.path.exists(dest_path):
                logger.debug(f"Overwriting existing file: {dest_file}")
            
            shutil.copy(source_path, dest_path)
            logger.debug(f"Copied {source_file} → {dest_file}")
            
        except KeyError as e:
            print_error(
                logger,
                f"File mapping missing required key {e}: {file_mapping}",
                error_type='invalid_config',
                context={'file_mapping': file_mapping}
            )
        except PermissionError as e:
            print_error(
                logger,
                f"Permission denied copying file '{source_file}' to {dest_folder}",
                e,
                error_type='permission_denied',
                context={'filename': source_file}
            )
        except Exception as e:
            print_error(logger, f"Could not copy file '{source_file}' to '{dest_path}'", e)
            
    logger.info(f"Copied {len(files_to_copy)} files successfully")
    
    # Copy OJS template
    try:
        if not os.path.exists(template_file):
            print_error(logger, f"OJS template file not found: {template_file}")
            
        new_ojs_file = os.path.join(
            tournament_folder,
            item[COL_SHORT_NAME],
            item[COL_OJS_FILENAME]
        )
        shutil.copy(template_file, new_ojs_file)
        logger.info(f"OJS template copied to: {new_ojs_file}")
    except Exception as e:
        print_error(logger, f"Could not copy OJS template '{template_file}' to '{new_ojs_file}'", e)


def generate_tournament_config(
    tournament: pd.Series,
    config: dict,
    dfAwardDef: pd.DataFrame,
    using_divisions: bool,
    tournament_folder: str,
    quiet: bool = False
) -> tuple[bool, str, list]:
    """Generate or update tournament_config.json file for a tournament.
    
    Args:
        tournament: Tournament row from dfTournaments
        config: Season configuration dictionary
        dfAwardDef: Award definitions DataFrame
        using_divisions: Whether this season uses divisions
        tournament_folder: Root tournament folder path
        quiet: If True, suppress info messages
        
    Returns:
        Tuple of (mismatch_detected, tournament_name, award_mismatches)
    """
    
    logger.info(f"Generating tournament config for {tournament[COL_SHORT_NAME]}")
    
    tourn_short = tournament[COL_SHORT_NAME]
    config_path = os.path.join(tournament_folder, tourn_short, 'tournament_config.json')
    
    # Track missing script tags for warning summary
    missing_tags_warnings = []
    
    # Check if config already exists (for second division)
    if os.path.exists(config_path):
        logger.debug(f"Config exists, will update: {config_path}")
        with open(config_path, 'r', encoding='utf-8') as f:
            existing_config = json.load(f)
    else:
        existing_config = None
    
    # Build INFO section (only on first pass)
    if existing_config is None:
        # Get tournament date
        tourn_date = tournament.get(COL_DATE, "")
        if pd.isna(tourn_date):
            tourn_date = ""
        else:
            tourn_date = str(tourn_date)
        
        # Determine OJS filenames for this tournament
        if using_divisions:
            # Get both division rows for this tournament
            ojs_filenames = []
            div1_filename = tournament.get(COL_OJS_FILENAME)
            if div1_filename and not pd.isna(div1_filename):
                ojs_filenames.append(div1_filename)
        else:
            ojs_filenames = [tournament[COL_OJS_FILENAME]]
        
        # Read dual_emcee flag from OJS file(s) - TRUE if ANY OJS has it set to TRUE
        dual_emcee = False
        
        info_section = {
            "season_name": config.get("season_name", ""),
            "season_year": config.get("season_yr", ""),
            "tournament_short_name": tourn_short,
            "tournament_long_name": tournament.get(COL_LONG_NAME, ""),
            "tournament_date": tourn_date,
            "using_divisions": using_divisions,
            "dual_emcee": dual_emcee,  # Will be updated below
            "ojs_filenames": ojs_filenames
        }
    else:
        # Keep existing INFO, but add this division's OJS filename if not present
        info_section = existing_config.get("INFO", {})
        current_ojs = tournament.get(COL_OJS_FILENAME)
        if current_ojs and not pd.isna(current_ojs):
            if current_ojs not in info_section.get("ojs_filenames", []):
                info_section.setdefault("ojs_filenames", []).append(current_ojs)
    
    # Read dual_emcee from current OJS file (cell F2 on Team and Program Information sheet)
    current_ojs_path = os.path.join(tournament_folder, tourn_short, tournament[COL_OJS_FILENAME])
    
    try:
        if os.path.exists(current_ojs_path):
            ojs_book = load_workbook(current_ojs_path, data_only=True)
            ws_tapi = ojs_book[SHEET_TEAM_INFO]
            dual_emcee_value = ws_tapi["F2"].value
            
            # Convert to boolean
            current_dual_emcee = False
            if isinstance(dual_emcee_value, bool):
                current_dual_emcee = dual_emcee_value
            elif isinstance(dual_emcee_value, str):
                current_dual_emcee = dual_emcee_value.upper() in ['TRUE', 'YES', '1']
            elif isinstance(dual_emcee_value, (int, float)):
                current_dual_emcee = bool(dual_emcee_value)
            
            # OR logic: if either OJS has TRUE, enable dual emcee
            if current_dual_emcee:
                info_section["dual_emcee"] = True
                logger.debug(f"Dual emcee enabled from {tournament[COL_OJS_FILENAME]}")
            
            ojs_book.close()
    except Exception as e:
        logger.debug(f"Could not read dual_emcee from {current_ojs_path}: {e}")
    
    # Build AWARDS section - use dictionary for easier merging
    awards_dict = {}
    award_mismatches = []
    mismatch_detected = False
    
    # Load existing awards if updating
    if existing_config:
        for award in existing_config.get("AWARDS", []):
            awards_dict[award["ID"]] = award
    
    # Get current division (if applicable)
    current_div = tournament.get(COL_DIVISION, "") if using_divisions else ""
    
    # Process all award columns from this tournament row
    award_columns = [col for col in tournament.index if col.startswith(AWARD_COLUMN_PREFIX_JUDGED) or col == AWARD_COLUMN_ROBOT_GAME]
    
    for award_col in award_columns:
        award_count = tournament.get(award_col, 0)
        
        # Skip if no allocation
        try:
            award_count = int(award_count) if not pd.isna(award_count) else 0
        except (ValueError, TypeError):
            award_count = 0
        
        if award_count == 0:
            continue
        
        # Get award metadata from AwardDef
        award_def_row = dfAwardDef[dfAwardDef[COL_COLUMN_NAME] == award_col]
        
        if award_def_row.empty:
            logger.warning(f"Award {award_col} not found in AwardDef table")
            continue
        
        award_name = award_def_row.iloc[0].get("Name", "")
        div_award_raw = award_def_row.iloc[0].get(COL_DIV_AWARD, False)
        
        # Get script tags from AwardDef
        script_tag_d1 = award_def_row.iloc[0].get(COL_SCRIPT_TAG_D1, "")
        script_tag_d2 = award_def_row.iloc[0].get(COL_SCRIPT_TAG_D2, "")
        script_tag_nodiv = award_def_row.iloc[0].get(COL_SCRIPT_TAG_NODIV, "")
        
        # Convert empty/NaN to empty string
        script_tag_d1 = "" if pd.isna(script_tag_d1) else str(script_tag_d1).strip()
        script_tag_d2 = "" if pd.isna(script_tag_d2) else str(script_tag_d2).strip()
        script_tag_nodiv = "" if pd.isna(script_tag_nodiv) else str(script_tag_nodiv).strip()
        
        # Get Labels from AwardDef
        label_columns = [col for col in dfAwardDef.columns if col.startswith(AWARD_LABEL_PREFIX)]
        labels = []
        for label_col in sorted(label_columns):  # Label1, Label2, Label3, etc.
            value = award_def_row.iloc[0].get(label_col)
            if pd.notna(value) and str(value).strip():
                labels.append(str(value).strip())
        
        # Convert DivAward to boolean
        if isinstance(div_award_raw, str):
            is_div_award = div_award_raw.upper() in ['TRUE', '1', 'YES']
        elif isinstance(div_award_raw, (int, float)):
            is_div_award = bool(div_award_raw)
        else:
            is_div_award = bool(div_award_raw)
        
        # Validate script tags based on award type
        if using_divisions and is_div_award:
            # Division award - should have D1 and D2 tags
            if not script_tag_d1 or not script_tag_d2:
                missing_tags_warnings.append(
                    f"Award {award_col} ({award_name}) is a division award but missing "
                    f"ScriptTag{'D1' if not script_tag_d1 else 'D2'}"
                )
        elif not is_div_award:
            # Tournament-level award or non-division tournament - should have NoDiv tag
            if not script_tag_nodiv:
                missing_tags_warnings.append(
                    f"Award {award_col} ({award_name}) missing ScriptTagNoDiv"
                )
        
        # Get or create award entry
        if award_col in awards_dict:
            award_entry = awards_dict[award_col]
        else:
            # Create new award entry
            award_entry = {
                "ID": award_col,
                "Name": award_name,
                "DivAwd": is_div_award,
                "Labels": labels
            }
            awards_dict[award_col] = award_entry
        
        # Add appropriate count fields based on division level
        if using_divisions and is_div_award:
            # Division-level award - add count for this division
            if current_div == "D1":
                award_entry["D1_count"] = award_count
            elif current_div == "D2":
                award_entry["D2_count"] = award_count
            
            # Add script tags for division awards (only if not empty)
            if script_tag_d1:
                award_entry["ScriptTagD1"] = script_tag_d1
            if script_tag_d2:
                award_entry["ScriptTagD2"] = script_tag_d2
        else:
            # Non-division tournament or tournament-level award
            award_entry["TournCount"] = award_count
            
            # Add script tag for non-division/tournament awards (only if not empty)
            if script_tag_nodiv:
                award_entry["ScriptTagNoDiv"] = script_tag_nodiv
    
    # Convert dictionary back to list
    awards_list = list(awards_dict.values())
    
    # Build final config structure
    final_config = {
        "INFO": info_section,
        "AWARDS": awards_list
    }
    
    # Write config file
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(final_config, f, indent=2)
        logger.info(f"Tournament config saved: {config_path}")
        if not quiet:
            logger.debug(f"Config contains {len(awards_list)} award(s)")
            logger.debug(f"Dual emcee: {info_section.get('dual_emcee', False)}")
    except Exception as e:
        logger.error(f"Failed to write tournament config: {e}")
    
    # Add missing tags warnings to the award_mismatches list for summary
    if missing_tags_warnings:
        award_mismatches.extend(missing_tags_warnings)
    
    return mismatch_detected, tourn_short, award_mismatches
