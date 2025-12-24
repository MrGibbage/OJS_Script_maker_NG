"""Logging configuration for the OJS tournament folder builder."""

import logging
import sys
import glob
from pathlib import Path
from datetime import datetime
from colorama import Fore, Style
from .user_feedback import get_error_recovery_suggestions
from typing import Optional
import os
import copy


class ColoredFormatter(logging.Formatter):
    """Custom formatter that adds color to console output."""
    
    COLORS = {
        'DEBUG': Fore.CYAN,
        'INFO': Fore.GREEN,
        'WARNING': Fore.YELLOW,
        'ERROR': Fore.RED,
        'CRITICAL': Fore.RED + Style.BRIGHT,
    }
    
    def format(self, record):
        # Make a copy of the record to avoid modifying the original
        record_copy = copy.copy(record)
        
        # Add color to the level name in the COPY only
        levelname = record_copy.levelname
        if levelname in self.COLORS:
            record_copy.levelname = f"{self.COLORS[levelname]}{levelname}{Style.RESET_ALL}"
        
        return super().format(record_copy)


def cleanup_old_logs(log_dir: str, log_prefix: str, current_log: str, keep_count: int = 1):
    """Delete old log files, keeping only the most recent ones.
    
    Args:
        log_dir: Directory containing log files
        log_prefix: Prefix of log files (e.g., "ojs_builder" or "ceremony_generator")
        current_log: Full path to the current log file (to exclude from deletion)
        keep_count: Number of most recent log files to keep (in addition to current)
    """
    try:
        # Normalize paths for comparison
        log_dir = os.path.normpath(log_dir)
        current_log = os.path.normpath(current_log)
        
        log_pattern = os.path.join(log_dir, f"{log_prefix}_*.log")
        log_files = glob.glob(log_pattern)
        
        # Normalize all found log file paths
        log_files = [os.path.normpath(f) for f in log_files]
        
        # Remove current log from the list
        log_files = [f for f in log_files if f != current_log]
        
        # If we want to keep only the current log (keep_count=1), delete ALL old logs
        if keep_count <= 1:
            files_to_delete = log_files
        else:
            # Keep (keep_count - 1) old logs
            if len(log_files) <= (keep_count - 1):
                return  # Nothing to delete
            
            # Sort by modification time (oldest first)
            log_files.sort(key=os.path.getmtime)
            
            # Delete all except the most recent (keep_count - 1)
            files_to_delete = log_files[:-(keep_count - 1)]
        
        # Delete the files
        for log_file in files_to_delete:
            try:
                os.remove(log_file)
            except (PermissionError, FileNotFoundError):
                pass  # Skip if locked or already gone
            except Exception:
                pass
                
    except Exception:
        pass  # Silently fail


def setup_logger(name: str = "ojs_builder", debug: bool = False, log_dir: Optional[str] = None) -> logging.Logger:
    """Configure and return a logger with both console and file handlers."""
    logger = logging.getLogger(name)
    
    # Set level
    level = logging.DEBUG if debug else logging.INFO
    logger.setLevel(level)
    
    # Avoid duplicate handlers if logger already configured
    if logger.handlers:
        return logger
    
    # File handler FIRST (gets record before console formatter modifies it)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = f"{name}_{timestamp}.log"
    
    if log_dir:
        log_path = Path(log_dir) / log_file
    else:
        log_path = Path.cwd() / log_file
    
    file_handler = logging.FileHandler(log_path, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    
    # Plain formatter (no color codes) for file
    plain_format = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(plain_format)
    logger.addHandler(file_handler)
    
    # Console handler SECOND (with color)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_format = ColoredFormatter('%(levelname)s: %(message)s')
    console_handler.setFormatter(console_format)
    logger.addHandler(console_handler)
    
    logger.info(f"Logging to file: {log_path}")
    
    # Cleanup old log files
    cleanup_old_logs(str(log_path.parent), name, str(log_path), keep_count=1)
    
    return logger


def print_error(logger: logging.Logger, errormsg: str, e: Optional[Exception] = None, error_type: Optional[str] = None, context: Optional[dict] = None) -> None:
    """Log an error message with recovery suggestions and exit the program.
    
    Args:
        logger: Logger instance to use
        errormsg: The error message to display
        e: Optional exception object to include in the output
        error_type: Type of error for recovery suggestions
        context: Additional context for recovery suggestions
    """
    if e:
        logger.error(f"{errormsg}\n{e}", exc_info=True)
    else:
        logger.error(errormsg)
    
    print(f"\n{Fore.RED}{'═' * 70}{Style.RESET_ALL}")
    print(f"{Fore.RED}ERROR: {errormsg}{Style.RESET_ALL}")
    if e:
        print(f"{Fore.RED}{e}{Style.RESET_ALL}")
    
    # Display recovery suggestions if error type provided
    if error_type and context:
        suggestions = get_error_recovery_suggestions(error_type, context)
        if suggestions:
            print(f"\n{Fore.YELLOW}Suggested Actions:{Style.RESET_ALL}")
            for i, suggestion in enumerate(suggestions, 1):
                print(f"  {i}. {suggestion}")
    
    print(f"{Fore.RED}{'═' * 70}{Style.RESET_ALL}\n")
    
    input("Press enter to quit...")
    sys.exit(1)
