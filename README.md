# Purpose:
Scripts and sample excel spreadsheets in here are used to prepare for VA-DC FLL tournaments

**2025-FLL-Qualifier-Tournaments.xlsx**: This is used with the build-tournament-folders.py script

**build-tournament-folders.py** (and build-tournament-folders.exe) does several things:
1. Creates folders for each tournament
2. Copies the template spreadsheet into the folder
3. Copies other needed files such as the closing ceremony script and templates
4. Populates the spreadsheet with the teams participating

**script-maker-mac-win.py** can be compiled (using pyinstaller) so it can run on Mac OS or Windows. This program is used by judge advisors to generate the closing ceremony scripts based on what is entered in the OJS.

**script_template.html.jinja** is the template file for creating the closing ceremony script

**2025-Qualifier-Template.xlsm** is the template file that is copied for each tournament OJS. It is a macro-enabled workbook, so be sure to enable macros when opening. There is a macro that will fill the tables with random scores which is useful for practicing and testing to see how the OJS works prior to tournament day. The worksheets use password protection to lock down cells that will not need additional entries. The Results and Ranking worksheet can be sorted by most columns using the OJS button menu.

# State Coordinator Instructions:
Clone the repository, or download these six files to a folder of your choice:
1. [2025-FLL-Qualifier-Tournaments.xlsx](2025/2025-FLL-Qualifier-Tournaments.xlsx)
2. [2025-Qualifier-Template.xlsm](2025/2025-Qualifier-Template.xlsm)
3. [build-tournament-folders.exe](2025/build-tournament-folders.exe)
4. [script_maker-win.exe](2025/script_maker-win.exe)
5. [script_maker-macos](2025/script_maker-macos)
6. [script_template.html.jinja](2025/script_template.html.jinja)

Or just grab [the files from the latest release ](https://github.com/MrGibbage/OJS_Script_maker_NG/releases).

Note that downloading exe files on Windows is tricky and requires some manual intervention on your part. I have instructions [here](DOWNLOADING.md).

Optional: If you want to be able to edit the python files, using VS code open the repository and build a python environment. Be sure to include the requirements from requirements.txt

If you cloned the repository, you will see a 2025 folder. Within that folder you will see all of the files for the 2025 season. If you just downloaded the files from here, you won't have a 2025 folder, but you will have all of the files you need. In either case, run the build-tournament-folders (either .py or .exe) program. You will see a new "tournaments" folder is created and folders within that for each tournament. Each tournament will have OJS files for each division at the tournament, the script_maker.exe file, and the script_template.html.jinja file which is the template for the closing ceremony script.

Within each tournament, update the OJS by entering the password (ask skip for instructions on how to do that, but it isn't particularly secure or hard to reverse engineer). After entering the password, use the "Update Teams" macro from the FLL toolbar. Optionally update the judging pod information. Then lock the worksheet from the FLL toolbar. Then use the check-setup program to validate the files. If you just want to check the setup of one tournament, you can enter the name here. The name is the same as the tournament folder name, such as "Norfolk". Case is important so pay attention to that. Also note that none of the tournament folder have spaces in them, but do have underscores and dashes as needed.

# OJS Script Maker NG

This repository contains tools for preparing scripts for judging events.

The Judging Timer app is located in the `judging_timer/` folder. Open `judging_timer/judging_timer.html` to run the app.

If you host this repository with GitHub Pages, the `judging_timer/` folder will be published and the app will be available at `https://<your-user>.github.io/<repo>/judging_timer/judging_timer.html`.
