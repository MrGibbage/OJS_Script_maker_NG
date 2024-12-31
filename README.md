Purpose:
Scripts and sample excel spreadsheets in here are used to prepare for VA-DC FLL tournaments

2024-FLL-Qualifier-Tournaments.xlsx: This is used with the build-tournament-folders.py script

build-tournament-folders.py (and build-tournament-folders.exe) does several things:
1. Creates folders for each tournament
2. Copies the template spreadsheet into the folder
3. Copies other needed files such as the closing ceremony script and templates
4. Populates the spreadsheet with the teams participating

check-setup.py (and check-setup.exe) will check a tournament (or all tournaments) to make sure the OJS files are named correctly and the files are ready to be sent to tournament directors and judge advisors

script-maker.py (and script-maker.exe) generates the closing ceremony scripts

script_template.html.jinja is the template file for creating the closing ceremony script

2024-Qualifier-Template.xlsm is the template file that is copied for each tournament OJS. It is a macro-enabled workbook, so be sure to enable macros when opening. Tournament Judge Advisors and Tournament Directors will not need to use the macros, so they can leave the macros disabled. The worksheets use password protection to lock down cells that will not need additional entries.
Macros are used to copy the team numbers and names to the other sheets
Protect and unprotext the workbooks
