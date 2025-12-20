"""Logging configuration for the OJS tournament folder builder."""

import logging
import sys
from pathlib import Path
from datetime import datetime
from colorama import Fore, Style
from .user_feedback import get_error_recovery_suggestions


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
        # Add color to the level name
        levelname = record.levelname
        if levelname in self.COLORS:
            record.levelname = f"{self.COLORS[levelname]}{levelname}{Style.RESET_ALL}"
        return super().format(record)


def setup_logger(name: str = "ojs_builder", log_dir: str | None = None, debug: bool = False) -> logging.Logger:
    """Configure and return a logger with both console and file handlers.
    
    Args:
        name: Name of the logger
        log_dir: Directory to store log files (defaults to script directory)
        debug: If True, set log level to DEBUG; otherwise INFO
        
    Returns:
        Configured logger instance
    """
    logger = logging.getLogger(name)
    
    # Set level
    level = logging.DEBUG if debug else logging.INFO
    logger.setLevel(level)
    
    # Avoid duplicate handlers if logger already configured
    if logger.handlers:
        return logger
    
    # Console handler with color
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_format = ColoredFormatter(
        '%(levelname)s: %(message)s'
    )
    console_handler.setFormatter(console_format)
    logger.addHandler(console_handler)
    
    # File handler (no color, more detail)
    if log_dir:
        log_path = Path(log_dir)
    else:
        log_path = Path.cwd()
    
    log_path.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_path / f"tournament_builder_{timestamp}.log"
    
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)  # Always log DEBUG to file
    file_format = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_format)
    logger.addHandler(file_handler)
    
    logger.info(f"Logging to file: {log_file}")
    
    return logger


def print_error(logger: logging.Logger, errormsg: str, e: Exception | None = None, error_type: str | None = None, context: dict | None = None) -> None:
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
