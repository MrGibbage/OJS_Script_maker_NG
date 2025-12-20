# OJS Tournament Folder Builder

Automated tool to prepare per-tournament folders and populate OJS (Online Judge System) spreadsheets for FIRST LEGO League tournaments.

## Installation

### Using uv (Recommended)

```bash
# Install uv if you haven't already
# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Create virtual environment and install dependencies
cd 2025-all
uv venv
uv pip install -e .
```

### Using pip (Alternative)

```bash
cd 2025-all
python -m venv .venv
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # macOS/Linux
pip install -e .
```

## Usage

### Quick Start (Default - Quiet Mode)

```bash
# Activate environment (if using uv)
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# Run the builder
python build-tournament-folders.py
```

You'll be prompted to select a tournament or press ENTER to build all.

### Command-Line Options

```bash
# Interactive mode (with prompts and validation summary)
python build-tournament-folders.py --interactive

# Verbose/debug mode
python build-tournament-folders.py --verbose

# Process specific tournament without prompts
python build-tournament-folders.py --tournament "Manassas_1"

# Combine options
python build-tournament-folders.py --verbose --tournament "ABC"

# Show help
python build-tournament-folders.py --help
```

### Available Flags

| Flag | Short | Description |
|------|-------|-------------|
| `--interactive` | `-i` | Show prompts, validation summary, and confirmations |
| `--verbose` | `-v` | Enable debug logging to console and file |
| `--tournament NAME` | `-t NAME` | Process only the specified tournament |
| `--skip-validation` | | Skip pre-flight checks (not recommended) |

## Modes

### Quiet Mode (Default)
- Minimal console output
- Logs everything to file
- Always prompts for tournament selection
- Best for regular use

### Interactive Mode (`--interactive`)
- Full validation summary display
- Confirmation prompts
- Progress indicators with status messages
- Best for troubleshooting or first-time setup

### Verbose Mode (`--verbose`)
- Debug-level logging to console and file
- Detailed operation information
- Shows table operations, file copies, etc.
- Best for debugging issues

## Configuration

Edit `season.json` to configure:

```json
{
  "season_yr": "2025",
  "season_name": "SUBMERGED",
  "filename": "2025-FLL-Qualifier-Tournaments.xlsx",
  "tournament_template": "2025-Qualifier-Template.xlsm",
  "copy_file_list": [
    "script_maker-win.exe",
    "script_maker-mac",
    "script_template.html.jinja",
    "instructions.pdf"
  ]
}
```

### Configuration Keys

- `season_yr`: Tournament season year
- `season_name`: FLL season theme name
- `filename`: Excel file with tournament list and assignments
- `tournament_template`: OJS template file to copy
- `copy_file_list`: Additional files to copy to each tournament folder

## Logs

Log files are automatically created with timestamps:

- **Location**: Same directory as the script
- **Format**: `tournament_builder_YYYYMMDD_HHMMSS.log`
- **Contents**: Full debug information (even in quiet mode)
- **Retention**: Manual cleanup (files are not auto-deleted)

### Reading Logs

Use logs to troubleshoot issues:

```bash
# View the most recent log
cat tournament_builder_*.log | tail -100

# Search for errors
grep ERROR tournament_builder_*.log

# Find specific tournament
grep "Manassas_1" tournament_builder_*.log
```

## Project Structure

```
2025-all/
├── build-tournament-folders.py    # Main script
├── modules/
│   ├── __init__.py
│   ├── constants.py               # Configuration constants
│   ├── logger.py                  # Logging setup
│   ├── file_operations.py         # File/folder operations
│   ├── excel_operations.py        # Excel table read/write
│   ├── worksheet_setup.py         # OJS worksheet configuration
│   └── user_feedback.py           # Progress tracking and validation
├── season.json                    # Season configuration
├── pyproject.toml                 # Project dependencies
└── tournaments/                   # Generated output (created by script)
    └── [tournament_name]/
        ├── [ojs_file].xlsm
        ├── script_maker-win.exe
        ├── script_maker-mac
        └── ...
```

## Troubleshooting

### Common Issues

**"Could not open tournament file"**
- Ensure Excel file is closed before running
- Check path in `season.json` is correct
- Verify file exists in the expected location

**"No teams assigned to [tournament], OJS file removed"**
- This is normal if a division has no teams
- The script automatically skips empty divisions
- Check the Assignments sheet if unexpected

**"Missing required columns in assignments"**
- Verify Assignments table has: `Team #`, `Team Name`, `Coach Name`
- Check for typos in column headers
- Ensure no extra spaces in column names

**Module import errors**
- Ensure virtual environment is activated
- Run `uv pip install -e .` to reinstall dependencies

### Getting Help

1. **Check the log file** - Contains detailed error information
2. **Run with `--verbose`** - Shows step-by-step execution
3. **Use `--interactive`** - See validation summary before processing
4. **Review error suggestions** - Script provides recovery steps for common issues

## Development

### Running Tests

```bash
# Install dev dependencies
uv pip install -e ".[dev]"
