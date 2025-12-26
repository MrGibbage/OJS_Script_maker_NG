# Purpose:
Scripts and sample excel spreadsheets in here are used to prepare for VA-DC FLL tournaments

**2025-FLL-Qualifier-Tournaments.xlsx**: This is used with the build-tournament-folders.py script

**build-tournament-folders.py** (and build-tournament-folders.exe) does several things:
1. Creates folders for each tournament
2. Copies the template spreadsheet into the folder
3. Copies other needed files such as the closing ceremony script template
4. Populates the spreadsheet with the teams participating
5. Generates tournament_config.json with tournament info and award allocations

**closing-ceremony-script-generator.py** validates OJS data and generates HTML closing ceremony scripts:
1. Validates all scores, awards, and team data
2. Collects team lists and award winners from OJS files
3. Renders ceremony script using Jinja2 template
4. Supports dual emcee mode with alternating color highlighting (enable via cell F2 in OJS)
5. Run from within tournament folder after OJS files are complete

**script-maker-mac-win.py** (legacy) can be compiled (using pyinstaller) to run on Mac OS or Windows. This program was used by judge advisors to generate closing ceremony scripts. It has been replaced by closing-ceremony-script-generator.py which provides better validation and features.

**script_template.html.jinja** is the template file for creating the closing ceremony script

**2025-Qualifier-Template.xlsm** is the template file that is copied for each tournament OJS. It is a macro-enabled workbook, so be sure to enable macros when opening. There is a macro that will fill the tables with random scores which is useful for practicing and testing to see how the OJS works prior to tournament day. The worksheets use password protection to lock down cells that will not need additional entries. The Results and Ranking worksheet can be sorted by most columns using the OJS button menu.

# State Coordinator Instructions:
Clone the repository, or download these files to a folder of your choice:
1. [2025-FLL-Qualifier-Tournaments.xlsx](2025-all/2025-FLL-Qualifier-Tournaments.xlsx)
2. [2025-Qualifier-Template.xlsm](2025-all/2025-Qualifier-Template.xlsm)
3. [build-tournament-folders.py](2025-all/build-tournament-folders.py) (or build-tournament-folders.exe)
4. [closing-ceremony-script-generator.py](2025-all/closing-ceremony-script-generator.py)
5. [script_template.html.jinja](2025-all/script_template.html.jinja)
6. [modules/](2025-all/modules/) (entire folder with all Python modules)
7. [season.json](2025-all/season.json)
8. [pyproject.toml](2025-all/pyproject.toml)

Or just grab [the files from the latest release](https://github.com/MrGibbage/OJS_Script_maker_NG/releases).

Note that downloading exe files on Windows is tricky and requires some manual intervention on your part. I have instructions [here](DOWNLOADING.md).

**Recommended**: Set up a Python environment using the instructions in [2025-all/README.md](2025-all/README.md) for the best experience with all features.

If you cloned the repository, you will see a 2025-all folder. Within that folder you will see all of the files for the 2025 season. If you just downloaded the files from here, you won't have a 2025-all folder, but you will have all of the files you need. 

## Workflow

1. **Build Tournament Folders**: Run build-tournament-folders.py to create tournament folders and populate OJS files with teams
2. **Complete OJS Files**: Update each tournament's OJS by:
   - Entering the password (ask coordinator for details)
   - Using the "Update Teams" macro from the FLL toolbar
   - Optionally updating judging pod information
   - Entering scores after tournament day
   - Selecting award winners
   - Marking advancing teams
3. **Optional - Enable Dual Emcee Mode**: Set cell F2 to TRUE in "Team and Program Information" sheet for alternating highlighting
4. **Generate Ceremony Script**: Navigate to tournament folder and run closing-ceremony-script-generator.py
5. **Validate**: Use check-setup program to validate files if needed

### Running the Ceremony Script Generator

```bash
# Navigate to a tournament folder
cd tournaments/Norfolk

# Run with verbose logging (recommended)
python ../../closing-ceremony-script-generator.py --verbose

# Or with debug logging for troubleshooting
python ../../closing-ceremony-script-generator.py --debug
```

The generator will:
- Validate all OJS data (scores, awards, advancing teams)
- Collect team lists and award winners
- Generate an HTML ceremony script file
- Display any warnings or errors

If you just want to check one tournament's setup, you can enter the name. The name is the same as the tournament folder name, such as "Norfolk". Case is important so pay attention to that. Also note that none of the tournament folders have spaces in them, but do have underscores and dashes as needed.

# OJS Script Maker NG

This repository contains tools for preparing scripts for judging events.

The Judging Timer app is located in the `judging_timer/` folder. Open `judging_timer/judging_timer.html` to run the app.

If you host this repository with GitHub Pages, the `judging_timer/` folder will be published and the app will be available at `https://<your-user>.github.io/<repo>/judging_timer/judging_timer.html`.
