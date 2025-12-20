"""User feedback and progress tracking utilities."""

import logging
from colorama import Fore, Style
from typing import List, Dict, Any


logger = logging.getLogger("ojs_builder")


class ProgressTracker:
    """Track and display progress for multi-step operations."""
    
    def __init__(self, total_steps: int, description: str = "Processing"):
        self.total_steps = total_steps
        self.current_step = 0
        self.description = description
        
    def update(self, step_name: str = "") -> None:
        """Update progress to next step."""
        self.current_step += 1
        percent = int((self.current_step / self.total_steps) * 100)
        bar_length = 30
        filled = int(bar_length * self.current_step / self.total_steps)
        bar = '█' * filled + '░' * (bar_length - filled)
        
        msg = f"{self.description}: [{bar}] {percent}% ({self.current_step}/{self.total_steps})"
        if step_name:
            msg += f" - {step_name}"
        
        print(f"\r{Fore.CYAN}{msg}{Style.RESET_ALL}", end='', flush=True)
        logger.debug(f"Progress: {self.current_step}/{self.total_steps} - {step_name}")
        
    def complete(self, message: str = "Complete!") -> None:
        """Mark progress as complete."""
        print(f"\r{Fore.GREEN}✓ {message}{' ' * 50}{Style.RESET_ALL}")
        logger.info(message)


class ValidationSummary:
    """Collect and display validation results before processing."""
    
    def __init__(self):
        self.errors: List[str] = []
        self.warnings: List[str] = []
        self.info: List[str] = []
        
    def add_error(self, message: str) -> None:
        """Add an error message."""
        self.errors.append(message)
        
    def add_warning(self, message: str) -> None:
        """Add a warning message."""
        self.warnings.append(message)
        
    def add_info(self, message: str) -> None:
        """Add an info message."""
        self.info.append(message)
        
    def has_errors(self) -> bool:
        """Check if there are any errors."""
        return len(self.errors) > 0
    
    def display(self) -> None:
        """Display the validation summary."""
        print("\n" + "=" * 60)
        print(f"{Fore.CYAN}VALIDATION SUMMARY{Style.RESET_ALL}".center(70))
        print("=" * 60)
        
        if self.errors:
            print(f"\n{Fore.RED}✗ ERRORS ({len(self.errors)}):{Style.RESET_ALL}")
            for i, err in enumerate(self.errors, 1):
                print(f"  {i}. {err}")
                
        if self.warnings:
            print(f"\n{Fore.YELLOW}⚠ WARNINGS ({len(self.warnings)}):{Style.RESET_ALL}")
            for i, warn in enumerate(self.warnings, 1):
                print(f"  {i}. {warn}")
                
        if self.info:
            print(f"\n{Fore.GREEN}ℹ INFO ({len(self.info)}):{Style.RESET_ALL}")
            for i, inf in enumerate(self.info, 1):
                print(f"  {i}. {inf}")
                
        print("=" * 60 + "\n")


def print_section_header(title: str) -> None:
    """Print a formatted section header."""
    print(f"\n{Fore.YELLOW}{'─' * 60}{Style.RESET_ALL}")
    print(f"{Fore.YELLOW}{title}{Style.RESET_ALL}")
    print(f"{Fore.YELLOW}{'─' * 60}{Style.RESET_ALL}\n")


def print_success(message: str) -> None:
    """Print a success message."""
    print(f"{Fore.GREEN}✓ {message}{Style.RESET_ALL}")
    logger.info(message)


def print_warning(message: str) -> None:
    """Print a warning message."""
    print(f"{Fore.YELLOW}⚠ {message}{Style.RESET_ALL}")
    logger.warning(message)


def print_info(message: str) -> None:
    """Print an info message."""
    print(f"{Fore.CYAN}ℹ {message}{Style.RESET_ALL}")
    logger.info(message)


def get_error_recovery_suggestions(error_type: str, context: Dict[str, Any]) -> List[str]:
    """Get recovery suggestions based on error type.
    
    Args:
        error_type: Type of error (e.g., 'missing_file', 'missing_sheet', 'permission_denied')
        context: Additional context about the error
        
    Returns:
        List of suggested recovery actions
    """
    suggestions = {
        'missing_config': [
            f"Ensure '{context.get('filename', 'season.json')}' exists in the script directory",
            "Check that the file name is spelled correctly",
            "Verify you're running the script from the correct directory",
        ],
        'invalid_json': [
            f"Open '{context.get('filename', 'season.json')}' in a text editor",
            "Check for syntax errors (missing commas, quotes, brackets)",
            "Use a JSON validator online to verify the file format",
            "Compare with a working example from a previous season",
        ],
        'missing_file': [
            f"Verify '{context.get('filename', 'the file')}' exists in {context.get('directory', 'the expected location')}",
            "Check the filename spelling and path in season.json",
            "Ensure the file hasn't been moved or renamed",
        ],
        'missing_template': [
            f"Locate the OJS template file: {context.get('filename', 'template file')}",
            "Copy it from a backup or previous season",
            "Update the 'tournament_template' path in season.json if the file is in a different location",
        ],
        'file_open': [
            f"Close '{context.get('filename', 'the file')}' in Excel or any other program",
            "Check Task Manager for hidden Excel processes and close them",
            "Restart your computer if the file remains locked",
        ],
        'permission_denied': [
            "Check file/folder permissions - you may need administrator rights",
            "Ensure the file is not marked as read-only",
            "Try running the script as administrator",
            "Close any programs that might be using the file",
        ],
        'missing_sheet': [
            f"Open '{context.get('workbook', 'the workbook')}' and verify sheet '{context.get('sheet_name', 'the sheet')}' exists",
            "Check for typos in sheet name",
            f"Available sheets: {', '.join(context.get('available_sheets', []))}",
            "Restore from backup if sheet was accidentally deleted",
        ],
        'missing_table': [
            f"Verify table '{context.get('table_name', 'the table')}' exists on sheet '{context.get('sheet_name', 'the sheet')}'",
            "In Excel: go to Table Design tab to see table names",
            f"Available tables on this sheet: {', '.join(context.get('available_tables', ['none']))}",
            "Recreate the table or restore from backup",
        ],
        'missing_columns': [
            f"Required columns: {', '.join(context.get('required', []))}",
            f"Missing columns: {', '.join(context.get('missing', []))}",
            "Check the spelling of column headers",
            "Ensure no extra spaces in column names",
            "Compare with template or previous year's file",
        ],
        'invalid_data': [
            f"Check the data in {context.get('location', 'the specified location')}",
            f"Expected: {context.get('expected', 'valid data')}",
            f"Found: {context.get('found', 'invalid data')}",
            "Fix the data and run the script again",
        ],
    }
    
    return suggestions.get(error_type, [
        "Check the error message above for details",
        "Review the log file for more information",
        "Ensure all required files are present and properly formatted",
        "Contact support if the issue persists",
    ])


def prompt_with_choices(prompt: str, choices: List[str], default: int = 0) -> str:
    """Prompt user with multiple choice options.
    
    Args:
        prompt: The question to ask
        choices: List of choice strings
        default: Index of default choice (0-based)
        
    Returns:
        The selected choice string
    """
    print(f"\n{Fore.CYAN}{prompt}{Style.RESET_ALL}")
    for i, choice in enumerate(choices, 1):
        default_marker = f" {Fore.GREEN}(default){Style.RESET_ALL}" if i - 1 == default else ""
        print(f"  {i}. {choice}{default_marker}")
    
    while True:
        response = input(f"\nEnter choice [1-{len(choices)}] or press ENTER for default: ").strip()
        
        if response == "":
            return choices[default]
        
        try:
            choice_num = int(response)
            if 1 <= choice_num <= len(choices):
                return choices[choice_num - 1]
            else:
                print(f"{Fore.RED}Please enter a number between 1 and {len(choices)}{Style.RESET_ALL}")
        except ValueError:
            print(f"{Fore.RED}Please enter a valid number{Style.RESET_ALL}")


def confirm_action(message: str, default: bool = True) -> bool:
    """Ask user to confirm an action.
    
    Args:
        message: Confirmation message
        default: Default response if user presses ENTER
        
    Returns:
        True if user confirms, False otherwise
    """
    default_str = "Y/n" if default else "y/N"
    response = input(f"{Fore.YELLOW}{message} [{default_str}]: {Style.RESET_ALL}").strip().lower()
    
    if response == "":
        return default
    
    return response in ['y', 'yes']
