"""File and folder operations for tournament setup."""

import os
import json
import shutil
import logging
from typing import Any
import pandas as pd

from .constants import COL_SHORT_NAME, COL_OJS_FILENAME, FOLDER_TOURNAMENTS
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
    extrafilelist: list[str]
) -> None:
    """Copy extra files and OJS template into tournament folder.
    
    Args:
        item: Tournament row with 'Short Name' and 'OJS_FileName'
        dir_path: Base directory containing tournaments
        template_file: Path to OJS template workbook
        extrafilelist: List of filenames to copy
        
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
                
            dest_folder = os.path.join(dir_path, FOLDER_TOURNAMENTS, item[COL_SHORT_NAME])
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
            dir_path,
            FOLDER_TOURNAMENTS,
            item[COL_SHORT_NAME],
            item[COL_OJS_FILENAME]
        )
        shutil.copy(template_file, new_ojs_file)
        logger.info(f"OJS template copied to: {new_ojs_file}")
    except Exception as e:
        print_error(logger, f"Could not copy OJS template '{template_file}' to '{new_ojs_file}'", e)
