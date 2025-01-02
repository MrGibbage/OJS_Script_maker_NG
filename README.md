# Purpose:
Scripts and sample excel spreadsheets in here are used to prepare for VA-DC FLL tournaments

**2024-FLL-Qualifier-Tournaments.xlsx**: This is used with the build-tournament-folders.py script

**build-tournament-folders.py** (and build-tournament-folders.exe) does several things:
1. Creates folders for each tournament
2. Copies the template spreadsheet into the folder
3. Copies other needed files such as the closing ceremony script and templates
4. Populates the spreadsheet with the teams participating

**check-setup.py** (and **check-setup.exe**) will check a tournament (or all tournaments) to make sure the OJS files are named correctly and the files are ready to be sent to tournament directors and judge advisors

**script-maker.py** (and **script-maker.exe**) generates the closing ceremony scripts

**script_template.html.jinja** is the template file for creating the closing ceremony script

**2024-Qualifier-Template.xlsm** is the template file that is copied for each tournament OJS. It is a macro-enabled workbook, so be sure to enable macros when opening. Tournament Judge Advisors and Tournament Directors will not need to use the macros on tournament day, so they can leave the macros disabled. There is however a macro that will fill the tables with random scores which does not require a password. This is useful for practicing and testing to see how the OJS works prior to tournament day. The worksheets use password protection to lock down cells that will not need additional entries.
Macros are used to copy the team numbers and names to the other sheets and to protect and unprotect the workbooks

# State Coordinator Instructions:
Clone the repository, or download these five files to a folder of your choice:
1. [2024-FLL-Qualifier-Tournaments.xlsx](2024/2024-FLL-Qualifier-Tournaments.xlsx)
2. [2024-Qualifier-Template.xlsm](2024/2024-Qualifier-Template.xlsm)
3. [build-tournament-folders.exe](2024/build-tournament-folders.exe)
4. [script_maker.exe](2024/script_maker.exe)
5. [script_template.html.jinja](2024/script_template.html.jinja)

Or just grab [the files from the releases ](https://github.com/MrGibbage/OJS_Script_maker_NG/releases).

Note that downloading exe files is tricky and requires some manual intervention on your part. I have instructions [here](DOWNLOADING.md).

Optional: If you want to be able to edit the python files, using VS code open the repository and build a python environment. Be sure to include the requirements from requirements.txt

Within the calendar year folder, such as 2024, run the build-tournament-folders (either .py or .exe) program

Within each tournament, update the OJS by entering the password (ask skip for instructions on how to do that, but it isn't particularly secure or hard to reverse engineer). After entering the password, use the "Update Teams" macro from the FLL toolbar. Optionally update the judging pod information. Then lock the worksheet from the FLL toolbar. Then use the check-setup program to validate the files.

Zip up each folder and send to the judge advisors

# Judge Advisor Instructions
Unzip the zip file to some location of your choice. On windows, common places are your desktop and your Documents folder.
You will either have one or two OJS files. One for each division. Open one of the OJS files. The spreadsheet has embedded macros, and your computer security will probably put up a warning. On tournament day, you will not need any macros, but for training, let's go ahead and fill in some random scores. Look at the toolbar at the top and choose the FLL Logo. One of the options is to fill the OJS with random scores. Don't worry, no matter how much you mess with the OJS files, you have a backup in your email that you can always go back to. Each of the worksheets should have all of the team numbers and names populated, along with their randomly generated scores. You will see that a lot of the cells are locked so that you can't edit them, but there are cells that are unlocked for the team scores.
