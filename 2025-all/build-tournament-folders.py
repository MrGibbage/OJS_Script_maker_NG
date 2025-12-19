"""Utility to prepare per-tournament folders and populate OJS spreadsheets.

This script reads a season manifest (`season.json`) and a master
`TournamentList` (or `DivTournamentList`) workbook to create a folder for
each tournament, copy template files, and populate the OJS (Online Judge
System) spreadsheet tables with team/award/meta information.
"""

import os
import sys
import warnings
from openpyxl import load_workbook
from colorama import init

# Import from modules
from modules.logger import setup_logger, print_error
from modules.constants import *
from modules.file_operations import load_json_without_notes, create_folder, copy_files
from modules.excel_operations import read_table_as_df, read_table_as_dict
from modules.worksheet_setup import (
    set_up_tapi_worksheet,
    set_up_award_worksheet,
    set_up_meta_worksheet,
    protect_worksheets,
    hide_worksheets,
)

# Suppress openpyxl warnings
warnings.simplefilter(action="ignore", category=UserWarning)

# Initialize colorama
init()

# Set up logger
logger = setup_logger("ojs_builder", debug=False)

def main():
    """Main execution function."""
    # Determine script directory
    if getattr(sys, "frozen", False):
        dir_path = os.path.dirname(sys.executable)
    elif __file__:
        dir_path = os.path.dirname(__file__)
    
    logger.info("=== OJS Tournament Folder Builder ===")
    logger.info(f"Working directory: {dir_path}")
    
    # Load configuration
    try:
        config = load_json_without_notes(os.path.join(dir_path, CONFIG_FILENAME))
    except FileNotFoundError:
        print_error(logger, f"Configuration file '{CONFIG_FILENAME}' not found in {dir_path}")
    except Exception as e:
        print_error(logger, f"Error loading configuration file '{CONFIG_FILENAME}'", e)

    # Validate required config keys
    required_config_keys = ["filename", "tournament_template", "copy_file_list"]
    missing_keys = [key for key in required_config_keys if key not in config]
    if missing_keys:
        print_error(logger, f"Missing required keys in {CONFIG_FILENAME}: {', '.join(missing_keys)}")

    tournament_file = os.path.join(dir_path, config["filename"])
    template_file = os.path.join(dir_path, config["tournament_template"])
    extrafilelist = config["copy_file_list"]

    # Validate required files exist
    if not os.path.exists(tournament_file):
        print_error(logger, f"Tournament file not found: {tournament_file}")
        
    if not os.path.exists(template_file):
        print_error(logger, f"Template file not found: {template_file}")

    # Check extra files
    for filename in extrafilelist:
        file_path = os.path.join(dir_path, filename)
        if os.path.exists(file_path):
            logger.info(f"✓ {filename}")
        else:
            logger.warning(f"✗ {filename} - MISSING!")

    # Load tournament data
    try:
        logger.info(f"Loading tournament data from {tournament_file}")
        _ = load_workbook(tournament_file, data_only=True)
    except Exception as e:
        print_error(
            logger,
            f"Could not open tournament file: {tournament_file}. "
            "Check to make sure it is not open in Excel.",
            e
        )

    # Read season info
    try:
        dictSeasonInfo = read_table_as_dict(tournament_file, "SeasonInfo", "SeasonInfo")
        using_divisions = dictSeasonInfo["Divisions"]
        logger.info(f"Using divisions: {using_divisions}")
    except Exception as e:
        print_error(logger, "Could not read the SeasonInfo table", e)

    # Read tournaments
    try:
        if using_divisions:
            dfTournaments = read_table_as_df(
                tournament_file, "DivTournaments", "DivTournamentList"
            ).fillna(0)
        else:
            dfTournaments = read_table_as_df(
                tournament_file, "Tournaments", "TournamentList"
            ).fillna(0)
        logger.info(f"Loaded {len(dfTournaments)} tournament(s)")
    except Exception as e:
        print_error(logger, "Could not read the tournament worksheet", e)

    # Read award definitions
    try:
        dfAwardDef = read_table_as_df(tournament_file, "AwardDef", "AwardDef").fillna(0)
        logger.info(f"Loaded {len(dfAwardDef)} award definitions")
    except Exception as e:
        print_error(logger, "Could not read the AwardDef worksheet", e)

    # Read assignments
    try:
        dfAssignments = read_table_as_df(tournament_file, "Assignments", "Assignments").fillna(0)
        tourn_array = dfTournaments[COL_SHORT_NAME].tolist()
        logger.info(f"Loaded {len(dfAssignments)} team assignments")
    except Exception as e:
        print_error(logger, "Could not read the assignments worksheet", e)

    # Tournament selection
    tourn = input("\nEnter tournament short name, or press ENTER for all tournaments: ")
    if tourn != "":
        if tourn in tourn_array:
            dfTournaments = dfTournaments.loc[dfTournaments[COL_SHORT_NAME] == tourn]
            logger.info(f"Building single tournament: {tourn}")
        else:
            print_error(
                logger,
                f"Tournament not found. Must be from this list: {tourn_array}"
            )

    # Process tournaments
    logger.info(f"\nProcessing {len(dfTournaments)} tournament(s)...")
    
    for index, row in dfTournaments.iterrows():
        logger.info("=" * 45)
        logger.info(f"  {row[COL_SHORT_NAME]} {row.get(COL_DIVISION, '')}  ".center(45))
        logger.info("=" * 45)
        
        # Create folder
        newpath = os.path.join(dir_path, FOLDER_TOURNAMENTS, row[COL_SHORT_NAME])
        create_folder(newpath)
        
        # Copy files
        logger.info("Copying files...")
        copy_files(row, dir_path, template_file, extrafilelist)

        # Process OJS file
        ojs_name = row.get(COL_OJS_FILENAME)
        if ojs_name is None or (isinstance(ojs_name, float) and pd.isna(ojs_name)):
            logger.warning(f"No OJS filename specified for {row[COL_SHORT_NAME]}, skipping")
            continue
            
        ojs_path = os.path.join(dir_path, FOLDER_TOURNAMENTS, row[COL_SHORT_NAME], ojs_name)
        logger.info(f"Processing OJS file: {ojs_name}")
        
        ojs_book = load_workbook(ojs_path, read_only=False, keep_vba=True)
        try:
            set_up_tapi_worksheet(row, ojs_book, dfAssignments, using_divisions)
            set_up_award_worksheet(row, ojs_book, dfAwardDef, using_divisions)
            set_up_meta_worksheet(row, ojs_book, config, dir_path, using_divisions)
            
            # Import remaining functions that weren't moved yet
            from modules.worksheet_setup import resize_worksheets, add_conditional_formats, copy_award_def
            
            add_conditional_formats(row, ojs_book)
            copy_award_def(row, ojs_book)
            hide_worksheets(row, ojs_book)
            resize_worksheets(row, ojs_book, dfAssignments, using_divisions)
            protect_worksheets(row, ojs_book)
            
        finally:
            ojs_book.save(ojs_path)
            ojs_book.close()
            logger.info(f"✓ Completed: {row[COL_SHORT_NAME]} {row.get(COL_DIVISION, '')}")

    logger.info("\n" + "=" * 45)
    logger.info("All tournaments processed successfully!")
    logger.info("=" * 45)


if __name__ == "__main__":
    main()