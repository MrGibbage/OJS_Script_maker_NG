# OJS Tournament Builder & Ceremony Script Generator

Automated tools to prepare per-tournament folders, populate OJS (Online Judge System) spreadsheets, and generate closing ceremony scripts for FIRST LEGO League tournaments.

## Features

- **Tournament Folder Builder**: Automatically creates tournament folders and populates OJS spreadsheets with team assignments
- **Closing Ceremony Script Generator**: Validates OJS data and generates HTML ceremony scripts with award winners
- **Dual Emcee Support**: Optional alternating color highlighting for two emcees reading the ceremony script
- **Conditional Formatting**: Visual feedback in OJS files for awards, ranks, and advancing teams
- **Comprehensive Validation**: Checks scores, ranges, and award allocations before ceremony script generation

## Installation

### Using uv (Recommended)

```bash
# Install uv if you haven't already
# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Create virtual environment and install all dependencies
cd 2025-all
uv sync
```

The `uv sync` command creates the virtual environment, installs all dependencies from `pyproject.toml`, and sets up the project in editable mode - all in one step!

### Using pip (Alternative)

```bash
cd 2025-all
python -m venv .venv
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # macOS/Linux
pip install -e .
```

## Usage

### Tournament Folder Builder

#### Quick Start (Default - Quiet Mode)

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

#### Command-Line Options

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

#### Available Flags

| Flag | Short | Description |
|------|-------|-------------|
| `--interactive` | `-i` | Show prompts, validation summary, and confirmations |
| `--verbose` | `-v` | Enable debug logging to console and file |
| `--tournament NAME` | `-t NAME` | Process only the specified tournament |
| `--skip-validation` | | Skip pre-flight checks (not recommended) |

### Closing Ceremony Script Generator

Run the ceremony script generator from within a tournament folder after OJS files are complete.

```bash
# Navigate to a tournament folder
cd tournaments/Norfolk

# Run the generator (verbose mode recommended)
python ../../closing-ceremony-script-generator.py --verbose

# Debug mode for troubleshooting
python ../../closing-ceremony-script-generator.py --debug
```

#### Features

- **Automatic Validation**: Checks all scores, awards, and team data before generating script
- **HTML Output**: Generates formatted ceremony script with proper headings and formatting
- **Dual Emcee Mode**: Enable by setting cell F2 to TRUE in the "Team and Program Information" sheet
  - When enabled, alternating lightblue/yellow highlighting on each paragraph
  - Helps two emcees track who reads next
  - Controlled at runtime - no need to regenerate config files
- **Award Integration**: Automatically populates award winners from OJS files
- **Division Support**: Handles both single and dual-division tournaments

#### Dual Emcee Highlighting

To enable dual emcee highlighting:
1. Open any OJS file for the tournament
2. Go to "Team and Program Information" worksheet
3. Set cell F2 to TRUE
4. Run the ceremony script generator

The generator checks all OJS files - if ANY file has F2=TRUE, highlighting is enabled.

#### Command-Line Options

| Flag | Short | Description |
|------|-------|-------------|
| `--verbose` | `-v` | Enable verbose logging (INFO level) |
| `--debug` | `-d` | Enable debug logging (DEBUG level, implies --verbose) |

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
  "tournament_folder": "C:/Users/username/Documents/tournaments",
  "copy_file_list_common": [
    {"source": "script_maker-win.exe", "dest": "script_maker-win.exe"},
    {"source": "script_maker-mac", "dest": "script_maker-mac"},
    {"source": "instructions.pdf", "dest": "instructions.pdf"}
  ],
  "copy_file_list_divisions_only": [
    {"source": "script_template-with-divisions.html.jinja", "dest": "script_template.html.jinja"},
    {"source": "summary_template-with-divisions.html.jinja", "dest": "summary_template.html.jinja"}
  ],
  "copy_file_list_no_divisions_only": [
    {"source": "script_template.html.jinja", "dest": "script_template.html.jinja"},
    {"source": "summary_template.html.jinja", "dest": "summary_template.html.jinja"}
  ]
}
```

### Configuration Keys

- `season_yr`: Tournament season year
- `season_name`: FLL season theme name
- `filename`: Excel file with tournament list and assignments
- `tournament_template`: OJS template file to copy
- `tournament_folder`: **Root folder where tournament subfolders will be created**
  - **IMPORTANT**: Use forward slashes `/` in paths (works on all platforms)
  - Will be created automatically if it doesn't exist
  - Must be an absolute path (full path from root)
  - Examples: 
    - Windows: `"C:/tournaments"` or `"C:/Users/username/Documents/tournaments"`
    - Mac: `"/Users/username/Documents/tournaments"`
    - Linux: `"/home/username/tournaments"`
- `copy_file_list_common`: Files always copied to every tournament folder
  - Each entry is a dict with `source` (filename in MAESTRO directory) and `dest` (filename in tournament folder)
  - Common files like executables and PDFs that don't vary by division setting
- `copy_file_list_divisions_only`: Files only copied when `using_divisions=True`
  - Typically includes division-specific templates (e.g., `script_template-with-divisions.html.jinja`)
  - Source and dest can differ to provide consistent naming in tournament folders
- `copy_file_list_no_divisions_only`: Files only copied when `using_divisions=False`
  - Typically includes non-division templates (e.g., `script_template.html.jinja`)
  - Allows TOAST to always use the same filenames regardless of division setting
  - ⚠️ **Don't use backslashes** `\` - they require escaping in JSON as `\\`
- `copy_file_list`: Additional files to copy to each tournament folder

## Logs

Log files are automatically created with timestamps in the script directory:

### Tournament Builder Logs
- **Format**: `tournament_builder_YYYYMMDD_HHMMSS.log`
- **Location**: `2025-all/` directory
- **Contents**: Full debug information (even in quiet mode)

### Ceremony Generator Logs
- **Format**: `ceremony_generator_YYYYMMDD_HHMMSS.log`
- **Location**: Tournament folder where you run the generator
- **Contents**: Validation results, data collection, and rendering details

### Log Retention
- Manual cleanup (files are not auto-deleted)
- Useful for troubleshooting and audit trails

### Reading Logs

Use logs to troubleshoot issues:

```bash
# View the most recent tournament builder log
cat tournament_builder_*.log | tail -100

# View the most recent ceremony generator log (from tournament folder)
cat ceremony_generator_*.log | tail -100

# Search for errors in any log
grep ERROR *.log

# Find specific tournament in builder logs
grep "Manassas_1" tournament_builder_*.log

# Check dual emcee status in ceremony logs
grep "dual_emcee" ceremony_generator_*.log
```

## Project Structure

```
2025-all/
├── build-tournament-folders.py          # Tournament folder builder
├── closing-ceremony-script-generator.py # Ceremony script generator
├── modules/
│   ├── __init__.py
│   ├── constants.py                     # Configuration constants
│   ├── logger.py                        # Logging setup
│   ├── file_operations.py               # File/folder operations & tournament config
│   ├── excel_operations.py              # Excel table read/write
│   ├── worksheet_setup.py               # OJS worksheet configuration & conditional formatting
│   ├── user_feedback.py                 # Progress tracking and validation
│   ├── ceremony_validator.py            # OJS data validation for ceremony scripts
│   ├── ceremony_data_collector.py       # Extract team/award data from OJS files
│   └── ceremony_renderer.py             # Jinja2 template rendering
├── script_template.html.jinja           # Ceremony script template
├── season.json                          # Season configuration
├── pyproject.toml                       # Project dependencies
└── [tournament_folder]/                 # Output location (specified in season.json)
    └── [tournament_name]/
        ├── [ojs_file].xlsm              # OJS spreadsheet with teams and scores
        ├── tournament_config.json        # Generated tournament configuration
        ├── [ceremony_script].html        # Generated ceremony script (after running generator)
        ├── script_template.html.jinja    # Ceremony template (copied here)
        ├── script_maker-win.exe
        ├── script_maker-mac
        └── ...
```

## Troubleshooting

### Tournament Builder Issues

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
- Run `uv sync` to reinstall dependencies

### Ceremony Script Generator Issues

**"Validation errors found"**
- Review the error messages - they indicate specific OJS data issues
- Check scores are within valid ranges (Innovation/Robot Design: 0-4, Core Values: 0-3)
- Verify all award selections match allocated counts
- Ensure Champion's Rank values are sequential starting from 1

**"Missing critical template variables"**
- Ensure tournament_config.json exists (run build-tournament-folders first)
- Check that OJS files have all required award selections
- Verify advancing teams are marked correctly

**"No highlighting in ceremony script"**
- Check cell F2 in "Team and Program Information" sheet is set to TRUE
- Verify the value is boolean TRUE (not text "TRUE")
- Try running with --debug to see which file enables dual emcee mode

**"Robot game awards not collecting"**
- Ensure "Robot Game Rank" column has sequential ranks (1, 2, 3...)
- Check "Max Robot Game Score" column has values
- Verify Team # and Team Name columns are populated

### Getting Help

1. **Check the log file** - Contains detailed error information
   - Tournament builder: `tournament_builder_YYYYMMDD_HHMMSS.log`
   - Ceremony generator: `ceremony_generator_YYYYMMDD_HHMMSS.log`
2. **Run with `--verbose` or `--debug`** - Shows step-by-step execution
3. **Use `--interactive`** - See validation summary before processing (builder only)
4. **Review error suggestions** - Scripts provide recovery steps for common issues

## Development

### Installing Dev Dependencies

```bash
# Install dev dependencies (pytest, black, ruff)
uv sync --extra dev
```
