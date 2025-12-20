"""File and folder operations for tournament setup."""

import os
import json
import shutil
import logging
from typing import Any
import pandas as pd

from .constants import (
    COL_SHORT_NAME, 
    COL_OJS_FILENAME, 
    COL_DATE, 
    COL_COLUMN_NAME, 
    COL_DIV_AWARD,
    AWARD_COLUMN_PREFIX_JUDGED
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
    extrafilelist: list[str],
    tournament_folder: str
) -> None:
    """Copy extra files and OJS template into tournament folder.
    
    Args:
        item: Tournament row with 'Short Name' and 'OJS_FileName'
        dir_path: Base directory containing source files
        template_file: Path to OJS template workbook
        extrafilelist: List of filenames to copy
        tournament_folder: Root folder where tournament subfolders are created
        
    Raises:
        Exits via print_error if any copy operation fails
    """
    if COL_SHORT_NAME not in item or not item[COL_SHORT_NAME]:
        print_error(logger, f"Tournament row missing required field '{COL_SHORT_NAME}'")
        
    if COL_OJS_FILENAME not in item or not item[COL_OJS_FILENAME]:
        print_error(logger, f"Tournament row missing required field '{COL_OJS_FILENAME}' for {item[COL_SHORT_NAME]}")
    
    logger.info(f"Copying files for tournament: {item[COL_SHORT_NAME]}")
    
    for filename in extrafilelist:
        try:
            source_path = os.path.join(dir_path, filename)
            if not os.path.exists(source_path):
                print_error(
                    logger,
                    f"Source file not found: {source_path}",
                    error_type='missing_file',
                    context={'filename': filename, 'directory': dir_path}
                )
                
            dest_folder = os.path.join(tournament_folder, item[COL_SHORT_NAME])
            shutil.copy(source_path, dest_folder)
            logger.debug(f"Copied {filename} to {dest_folder}")
        except PermissionError as e:
            print_error(
                logger,
                f"Permission denied copying file '{filename}' to {dest_folder}",
                e,
                error_type='permission_denied',
                context={'filename': filename}
            )
        except Exception as e:
            print_error(logger, f"Could not copy file '{filename}' to {dest_folder}", e)
            
    logger.info("Extra files copied successfully")
    
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
    tournament_row: pd.Series,
    config: dict,
    dfAwardDef: pd.DataFrame,
    using_divisions: bool,
    tournament_folder: str,
    quiet: bool = False
) -> tuple[bool, str, list[str]]:
    """Generate or update tournament_config.json file for a tournament.
    
    Creates a JSON configuration file in the tournament folder with tournament
    metadata and settings. For tournaments with divisions, updates existing file
    by appending to ojs_filenames array. Validates that tournament-level award
    counts match between divisions.
    
    Args:
        tournament_row: Row from dfTournaments with tournament data
        config: Season configuration dictionary (from season.json)
        dfAwardDef: Award definitions dataframe
        using_divisions: Whether season is using divisions
        tournament_folder: Root folder where tournament subfolders are created
        quiet: If True, suppress info messages
        
    Returns:
        Tuple of (mismatch_detected: bool, tournament_name: str, award_count_mismatches: list[str])
        mismatch_detected is True if JSON existed with using_divisions=False
        award_count_mismatches contains descriptions of any award count mismatches found
        
    Raises:
        Exits via print_error if critical errors occur
    """
    from colorama import Fore, Style
    
    tournament_short_name = tournament_row[COL_SHORT_NAME]
    tournament_name = tournament_short_name
    config_file_path = os.path.join(tournament_folder, tournament_short_name, "tournament_config.json")
    
    logger.debug(f"Generating config for: {tournament_short_name}")
    
    # Check if config file already exists
    existing_config = None
    mismatch_detected = False
    award_count_mismatches = []
    
    if os.path.exists(config_file_path):
        try:
            with open(config_file_path, "r", encoding="utf-8") as f:
                existing_config = json.load(f)
            logger.debug(f"Found existing config file: {config_file_path}")
        except json.JSONDecodeError as e:
            print_error(
                logger,
                f"Existing tournament_config.json is invalid JSON: {config_file_path}",
                e,
                error_type='invalid_json',
                context={'filename': 'tournament_config.json'}
            )
        except Exception as e:
            print_error(
                logger,
                f"Could not read existing tournament_config.json: {config_file_path}",
                e
            )
    
    # Handle division mismatch scenario
    if existing_config:
        if not existing_config.get("using_divisions", False):
            # Mismatch: JSON says no divisions, but we're seeing it again (must be division 2)
            mismatch_detected = True
            warning_msg = (
                f"⚠ DIVISION MISMATCH for {tournament_short_name}: "
                f"tournament_config.json has using_divisions=false, but tournament appears to use divisions. "
                f"Changing to using_divisions=true. Check season workbook settings if this is unexpected."
            )
            print(f"\n{Fore.YELLOW}{warning_msg}{Style.RESET_ALL}\n")
            logger.warning(warning_msg)
            
            # Update existing config: change using_divisions and append filename
            existing_config["using_divisions"] = True
            existing_config["ojs_filenames"].append(tournament_row[COL_OJS_FILENAME])
            
            try:
                with open(config_file_path, "w", encoding="utf-8") as f:
                    json.dump(existing_config, f, indent=2)
                logger.info(f"✓ Updated tournament_config.json (mismatch corrected): {config_file_path}")
                if not quiet:
                    print(f"{Fore.GREEN}✓ Updated tournament config (division mismatch corrected){Style.RESET_ALL}")
            except Exception as e:
                print_error(
                    logger,
                    f"Could not write updated tournament_config.json: {config_file_path}",
                    e,
                    error_type='file_write',
                    context={'filename': 'tournament_config.json'}
                )
            
            return (mismatch_detected, tournament_name, award_count_mismatches)
        else:
            # Normal division 2 scenario: append filename and validate award counts
            
            # For using_divisions tournaments, validate tournament-level award counts match
            if using_divisions and "tournament_award_counts" in existing_config:
                stored_counts = existing_config["tournament_award_counts"]
                
                for award_col, div1_count in stored_counts.items():
                    # Get the count from current row (division 2)
                    current_count = tournament_row.get(award_col, 0)
                    if pd.isna(current_count):
                        current_count = 0
                    else:
                        current_count = int(current_count)
                    
                    # Compare counts
                    if current_count != div1_count:
                        mismatch_msg = f"{award_col}: Div1={div1_count}, Div2={current_count}"
                        award_count_mismatches.append(mismatch_msg)
                        warning_msg = (
                            f"⚠ AWARD COUNT MISMATCH for {tournament_short_name}: "
                            f"{award_col} has Div1={div1_count} but Div2={current_count}. "
                            f"Tournament-level awards should have the same count. Check season workbook."
                        )
                        print(f"\n{Fore.YELLOW}{warning_msg}{Style.RESET_ALL}\n")
                        logger.warning(warning_msg)
            
            existing_config["ojs_filenames"].append(tournament_row[COL_OJS_FILENAME])
            
            try:
                with open(config_file_path, "w", encoding="utf-8") as f:
                    json.dump(existing_config, f, indent=2)
                logger.info(f"✓ Updated tournament_config.json (added division): {config_file_path}")
                if not quiet:
                    print(f"{Fore.GREEN}✓ Updated tournament config (added division){Style.RESET_ALL}")
            except Exception as e:
                print_error(
                    logger,
                    f"Could not write updated tournament_config.json: {config_file_path}",
                    e,
                    error_type='file_write',
                    context={'filename': 'tournament_config.json'}
                )
            
            return (mismatch_detected, tournament_name, award_count_mismatches)
    
    # Create new config file
    logger.debug("Creating new tournament_config.json")
    
    # Extract tournament data
    tournament_long_name = tournament_row.get("Long Name", tournament_short_name)
    tournament_date = tournament_row.get(COL_DATE, "")
    
    # Build award list: columns starting with J_AWD or P_AWD where value > 0
    tournament_award_list = []
    tournament_award_counts = {}  # Store counts for division validation
    
    for col in tournament_row.index:
        # Check if column is an award column
        if col.startswith(AWARD_COLUMN_PREFIX_JUDGED) or col.startswith("P_AWD"):
            # Check if this tournament has this award (value > 0)
            value = tournament_row[col]
            if pd.notna(value) and value > 0:
                # If using divisions, filter by DivAward = FALSE
                if using_divisions:
                    # Find matching row in dfAwardDef
                    matching_awards = dfAwardDef[dfAwardDef[COL_COLUMN_NAME] == col]
                    if not matching_awards.empty:
                        div_award_value = matching_awards.iloc[0][COL_DIV_AWARD]
                        # Only include if DivAward is FALSE (could be boolean or string)
                        if isinstance(div_award_value, bool):
                            if not div_award_value:
                                tournament_award_list.append(col)
                                tournament_award_counts[col] = int(value)
                        elif isinstance(div_award_value, str):
                            if div_award_value.upper() == "FALSE":
                                tournament_award_list.append(col)
                                tournament_award_counts[col] = int(value)
                        elif div_award_value == 0:  # Sometimes FALSE is stored as 0
                            tournament_award_list.append(col)
                            tournament_award_counts[col] = int(value)
                    else:
                        logger.warning(f"Award column '{col}' not found in AwardDef for {tournament_short_name}")
                else:
                    # Not using divisions: all awards are tournament-level
                    tournament_award_list.append(col)
                    tournament_award_counts[col] = int(value)
    
    logger.debug(f"Found {len(tournament_award_list)} tournament-level awards")
    
    # Build config dictionary
    config_data = {
        "season_name": config.get("season_name", ""),
        "season_year": config.get("season_yr", ""),
        "tournament_short_name": tournament_short_name,
        "tournament_long_name": tournament_long_name,
        "tournament_date": str(tournament_date) if pd.notna(tournament_date) else "",
        "using_divisions": using_divisions,
        "ojs_filenames": [tournament_row[COL_OJS_FILENAME]],
        "tournament_award_counts": tournament_award_counts
    }
    
    # Write config file
    try:
        with open(config_file_path, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=2)
        logger.info(f"✓ Created tournament_config.json: {config_file_path}")
        if not quiet:
            print(f"{Fore.GREEN}✓ Created tournament config file{Style.RESET_ALL}")
    except Exception as e:
        print_error(
            logger,
            f"Could not write tournament_config.json: {config_file_path}",
            e,
            error_type='file_write',
            context={'filename': 'tournament_config.json'}
        )
    
    return (mismatch_detected, tournament_name, award_count_mismatches)
