"""Validation logic for OJS spreadsheets in closing ceremony script generation."""

import logging
import pandas as pd
from openpyxl.workbook import Workbook
from typing import List, Tuple

from .constants import (
    SHEET_ROBOT_GAME, SHEET_INNOVATION, SHEET_ROBOT_DESIGN, SHEET_CORE_VALUES,
    TABLE_ROBOT_GAME, TABLE_INNOVATION, TABLE_ROBOT_DESIGN, TABLE_CORE_VALUES,
    COL_TEAM_NUMBER
)
from .excel_operations import read_table_as_df

logger = logging.getLogger("ceremony_generator")


class ValidationError:
    """Represents a validation error with context."""
    
    def __init__(self, severity: str, sheet: str, message: str):
        self.severity = severity  # "ERROR" or "WARNING"
        self.sheet = sheet
        self.message = message
    
    def __str__(self):
        return f"[{self.severity}] {self.sheet}: {self.message}"


class OJSValidator:
    """Validates OJS workbook data for ceremony script generation."""
    
    def __init__(self):
        self.errors: List[ValidationError] = []
        self.warnings: List[ValidationError] = []
    
    def add_error(self, sheet: str, message: str):
        """Add a validation error."""
        error = ValidationError("ERROR", sheet, message)
        self.errors.append(error)
        logger.error(str(error))
    
    def add_warning(self, sheet: str, message: str):
        """Add a validation warning."""
        warning = ValidationError("WARNING", sheet, message)
        self.warnings.append(warning)
        logger.warning(str(warning))
    
    def has_errors(self) -> bool:
        """Check if there are any errors."""
        return len(self.errors) > 0
    
    def validate_robot_game_scores(self, ojs_path: str, division: str = "") -> bool:
        """Validate Robot Game scores are within valid range and no blanks.
        
        Args:
            ojs_path: Path to OJS workbook
            division: Division label for error messages
            
        Returns:
            True if validation passed, False otherwise
        """
        logger.info(f"Validating Robot Game scores{' for ' + division if division else ''}")
        
        try:
            df = read_table_as_df(ojs_path, SHEET_ROBOT_GAME, TABLE_ROBOT_GAME)
        except Exception as e:
            self.add_error(SHEET_ROBOT_GAME, f"Could not read table: {e}")
            return False
        
        score_columns = ["Robot Game 1 Score", "Robot Game 2 Score", "Robot Game 3 Score"]
        
        for col in score_columns:
            if col not in df.columns:
                self.add_error(SHEET_ROBOT_GAME, f"Missing column: {col}")
                continue
            
            # Check for blanks
            blank_count = df[col].isna().sum()
            if blank_count > 0:
                self.add_error(SHEET_ROBOT_GAME, f"{col} has {blank_count} blank cell(s)")
            
            # Check range (0-545)
            valid_scores = df[col].dropna()
            out_of_range = valid_scores[(valid_scores < 0) | (valid_scores > 545)]
            if len(out_of_range) > 0:
                self.add_error(
                    SHEET_ROBOT_GAME,
                    f"{col} has {len(out_of_range)} score(s) outside valid range (0-545)"
                )
        
        return not self.has_errors()
    
    def validate_rubric_scores(
        self, 
        ojs_path: str, 
        sheet_name: str, 
        table_name: str,
        columns: List[str],
        division: str = ""
    ) -> bool:
        """Validate rubric scores are 0-5 and no blanks.
        
        Args:
            ojs_path: Path to OJS workbook
            sheet_name: Name of worksheet
            table_name: Name of table
            columns: List of column names to validate
            division: Division label for error messages
            
        Returns:
            True if validation passed, False otherwise
        """
        logger.info(f"Validating {sheet_name} scores{' for ' + division if division else ''}")
        
        try:
            df = read_table_as_df(ojs_path, sheet_name, table_name)
        except Exception as e:
            self.add_error(sheet_name, f"Could not read table: {e}")
            return False
        
        for col in columns:
            if col not in df.columns:
                self.add_error(sheet_name, f"Missing column: {col}")
                continue
            
            # Check for blanks
            blank_count = df[col].isna().sum()
            if blank_count > 0:
                self.add_error(sheet_name, f"{col} has {blank_count} blank cell(s)")
            
            # Check range (0-5)
            valid_scores = df[col].dropna()
            out_of_range = valid_scores[(valid_scores < 0) | (valid_scores > 5)]
            if len(out_of_range) > 0:
                self.add_error(
                    sheet_name,
                    f"{col} has {len(out_of_range)} score(s) outside valid range (0-5)"
                )
        
        return not self.has_errors()
    
    def validate_core_values_scores(self, ojs_path: str, division: str = "") -> bool:
        """Validate Core Values scores are in [0, 2, 3, 4] and no blanks.
        
        Args:
            ojs_path: Path to OJS workbook
            division: Division label for error messages
            
        Returns:
            True if validation passed, False otherwise
        """
        logger.info(f"Validating Core Values scores{' for ' + division if division else ''}")
        
        try:
            df = read_table_as_df(ojs_path, SHEET_CORE_VALUES, TABLE_CORE_VALUES)
        except Exception as e:
            self.add_error(SHEET_CORE_VALUES, f"Could not read table: {e}")
            return False
        
        cv_columns = [
            "Gracious Professionalism 1",
            "Gracious Professionalism 2",
            "Gracious Professionalism 3"
        ]
        
        valid_values = {0, 2, 3, 4}
        
        for col in cv_columns:
            if col not in df.columns:
                self.add_error(SHEET_CORE_VALUES, f"Missing column: {col}")
                continue
            
            # Check for blanks
            blank_count = df[col].isna().sum()
            if blank_count > 0:
                self.add_error(SHEET_CORE_VALUES, f"{col} has {blank_count} blank cell(s)")
            
            # Check valid values
            scores = df[col].dropna()
            invalid = scores[~scores.isin(valid_values)]
            if len(invalid) > 0:
                self.add_error(
                    SHEET_CORE_VALUES,
                    f"{col} has {len(invalid)} invalid score(s). Must be one of: {sorted(valid_values)}"
                )
        
        return not self.has_errors()
    
    def validate_all_sheets(self, ojs_path: str, division: str = "") -> bool:
        """Run all validations on an OJS workbook.
        
        Args:
            ojs_path: Path to OJS workbook
            division: Division label for error messages
            
        Returns:
            True if all validations passed, False otherwise
        """
        logger.info(f"Starting complete validation{' for ' + division if division else ''}")
        
        # Robot Game
        self.validate_robot_game_scores(ojs_path, division)
        
        # Innovation Project
        ip_columns = [
            "Identify - Define", "Identify - Research (CV)", "Design - Plan",
            "Design - Teamwork (CV)", "Create - Innovation (CV)", "Create - Model",
            "Iterate - Sharing", "Iterate - Improvement", "Communicate - Impact (CV)",
            "Communicate - Fun (CV)"
        ]
        self.validate_rubric_scores(
            ojs_path, SHEET_INNOVATION, TABLE_INNOVATION, ip_columns, division
        )
        
        # Robot Design
        rd_columns = [
            "Identify - Strategy", "Identify - Research (CV)", "Design - Ideas (CV)",
            "Design - Building/Coding", "Create - Attachments", "Create - Code/ Sensors",
            "Iterate - Testing", "Iterate - Improvements (CV)", "Communicate - Impact (CV)",
            "Communicate - Fun (CV)"
        ]
        self.validate_rubric_scores(
            ojs_path, SHEET_ROBOT_DESIGN, TABLE_ROBOT_DESIGN, rd_columns, division
        )
        
        # Core Values
        self.validate_core_values_scores(ojs_path, division)
        
        return not self.has_errors()
