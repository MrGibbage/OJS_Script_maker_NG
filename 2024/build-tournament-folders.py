# To create an executable file, run
# .venv\Scripts\pyinstaller.exe -F 2024\build-tournament-folders.py
# Then copy the build-tournament-folders.exe file from dist to 2024
#
import os, sys, re

# pip install openpyxl
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, get_column_letter, coordinate_from_string
from openpyxl.worksheet.table import Table
# pip install print-color
from print_color import print

from typing import Dict

import shutil

import numpy as np

import pandas as pd

import warnings

# If this line isn't here, we will get a UserWarning when we run the program, alerting us that the
# conditional formatting in the spreadsheets will not be preserved when we copy the data.
# We don't need conditional formatting within the data manipulation, so it isn't a big deal
# UserWarning: Data Validation extension is not supported and will be removed
# https://stackoverflow.com/questions/53965596/python-3-openpyxl-userwarning-data-validation-extension-not-supported
warnings.simplefilter(action='ignore', category=UserWarning)

cwd: str = os.getcwd()
dir_path = os.path.dirname(os.path.realpath(__file__))
current_year: str = '2024'
tournament_file: str = dir_path + '\\2024-FLL-Qualifier-Tournaments.xlsx'
template_file: str = dir_path + '\\2024-Qualifier-Template.xlsm'

# from https://stackoverflow.com/questions/56923379/how-to-read-an-existing-worksheet-table-with-openpyxl
def read_excel_table(sheet, table_name):
    """
    This function will read an Excel table
    and return a tuple of columns and data

    This function assumes that tables have column headers
    :param sheet: the sheet
    :param table_name: the name of the table
    :return: columns (list) and data (dict)
    """
    table = sheet.tables[table_name]
    table_range = table.ref

    table_head = sheet[table_range][0]
    table_data = sheet[table_range][1:]

    columns = [column.value for column in table_head]
    data = {column: [] for column in columns}

    for row in table_data:
        row_val = [cell.value for cell in row]
        for key, val in zip(columns, row_val):
            data[key].append(val)

    return columns, data

# creates the subfolders for each tournament
def create_folder(newpath):
    if os.path.exists(newpath):
        return
    try:
        os.makedirs(newpath)
        print(
            f'Created folder: {newpath}',
            tag="ok",
            tag_color="white",
            color="white",
        )
    except Exception as e:
        print(
            f'Could not create directory: {newpath}',
            tag="error",
            tag_color="red",
            color="red",
        )

# copies the files into the tournament folders
def copy_files(item:pd.Series):
    for filename in extrafilelist:
        try:
            newpath = dir_path + "\\tournaments\\" + item['Short Name'] + "\\"
            shutil.copy(filename, newpath)
        except Exception as e:
            print(
                f'Could not copy file: {filename} to {newpath}\n{e}',
                tag="error",
                tag_color="red",
                color="red",
            )
    # next copy the template files
    ojsfilelist.clear()
    if item['D1_OJS'] is not None:
        ojsfilelist.append(item['D1_OJS'])
    if item['D2_OJS'] is not None:
        ojsfilelist.append(item['D2_OJS'])
    for filename in ojsfilelist:
        try:
            shutil.copy(template_file, dir_path + "\\tournaments\\" + item["Short Name"] + "\\" + filename)
            # print(f'Copied {template_file} to {startpath + item["Short Name"] + "\\" + filename}')
        except Exception as  e:
            print(
                f'Could not copy OJS file: {template_file} to {dir_path + "\\tournaments\\" + item["Short Name"] + "\\" + filename}',
                tag="error",
                tag_color="red",
                color="red",
            )

# edits the OJS spreadsheets with the correct tournament information
def set_up_tapi_worksheet(tournament:pd.Series):
    # open the OJS workbook
    for d in ["D1", "D2"]:
        print(f"Setting up Tournament Team and Program Information for {tournament["Short Name"]} {d}")
        divassignees: pd.DataFrame = dfAssignments[(dfAssignments["Tournament"] == tournament["Short Name"]) & (dfAssignments["Div"] == d)]
        print(divassignees)
        if len(divassignees.index) > 0:
            try:
                # print(divassignees)
                ojsfile = dir_path + "\\tournaments\\" + tournament["Short Name"] + "\\" + tournament[d + "_OJS"]
                ojs_book = load_workbook(ojsfile, read_only=False, keep_vba=True)
                ws = ojs_book['Team and Program Information']
                table: Table = ws.tables["OfficialTeamList"]
                table_range: str = table.ref # should return a string C2:F3
                start_cell = table_range.split(':')[0] # should return 'C2'
                coords = coordinate_from_string(start_cell)
                start_col = column_index_from_string(coords[0])
                start_index = divassignees.index[0]
                table.ref = table.ref[:-1] + str(len(divassignees) + 2)
                for i, row in divassignees.iterrows():
                    # cell = f'{get_column_letter(start_col)}{i + 3 - start_index}'
                    # print(f'Cell: {cell}, setting value {row['Team #']}')
                    # ws[cell] = row['Team #']
                    ws.cell(row=i + 3 - start_index, column=start_col).value = row['Team #']
                    ws.cell(row=i + 3 - start_index, column=start_col + 1).value = row['Team Name']
                    ws.cell(row=i + 3 - start_index, column=start_col + 2).value = row['Coach Name']

                # Save the workbook
                # print('saving')
                ojs_book.save(ojsfile)
                # ojs_book.close()
            except Exception as e:
                print(
                    f'There was an error: {e}',
                    tag="error",
                    tag_color="red",
                    color="red",
                )

def set_up_award_worksheet(tournament:pd.Series, judge_awards: int):
    for d in ["D1", "D2"]:
        print(f"Setting up Tournament Awards for {tournament["Short Name"]} {d}")
        divawards: pd.DataFrame = dfAwards[(dfAwards["Tournament"] == tournament["Short Name"]) & (dfAwards["Div"] == d)]
        divawards = divawards.transpose()
        divawards = divawards.reset_index()
        divawards.columns = ['Award', 'Count']
        divawards.drop(0, inplace=True)
        divawards.drop(1, inplace=True)
        divawards.drop(divawards[divawards["Award"] == "ADV"].index, inplace=True)
        print(divawards)
        rg_awards: pd.DataFrame = divawards[(divawards["Award"].str.startswith("RG")) & (divawards['Count'] == 1)]
        print(rg_awards)

        # Robot Game
        if tournament[d + "_OJS"] is None:
            print(f"Nothing for tournament[{d}]")
            continue

        if len(rg_awards.index) > 0:
            print("Looking for robot game awards")
            ojsfile = dir_path + "\\tournaments\\" + tournament["Short Name"] + "\\" + tournament[d + "_OJS"]
            ojs_book = load_workbook(ojsfile, read_only=False, keep_vba=True)
            ws = ojs_book['AwardList']
            table: Table = ws.tables["RobotGameAwards"]
            table_range: str = table.ref 
            print(f"table_range: {table_range}")
            start_cell = table_range.split(':')[0]
            coords = coordinate_from_string(start_cell)
            start_col = column_index_from_string(coords[0])
            # start_index = divassignees.index[0]
            table.ref = table.ref[:-1] + str(len(rg_awards) + 1)
            next_row = 2
            for i, row in rg_awards.iterrows():
                print(f'New row, looking for RG {row}')
                award = row['Award'].replace("RG", "Robot Game ")
                award = award.replace("1", "1st Place").replace("2", "2nd Place").replace("3", "3rd Place")
                if award[:2] == "Ro" and row['Count'] == 1:
                    ws.cell(row=next_row, column=start_col).value = award
                    print(f"Adding the robot game award: {award} to cell(row = {next_row}, column = {start_col})")
                    next_row = next_row + 1

        # Other judged awards
        divawards.drop(divawards[divawards["Award"] == "RG1"].index, inplace=True)
        divawards.drop(divawards[divawards["Award"] == "RG2"].index, inplace=True)
        divawards.drop(divawards[divawards["Award"] == "RG3"].index, inplace=True)
        divawards = divawards[(divawards["Count"] > 0)]
        divawards = divawards.reset_index(drop=True)
        print(divawards)

        if len(divawards.index) > 0:
            # ojsfile = dir_path + "\\tournaments\\" + tournament["Short Name"] + "\\" + tournament[d + "_OJS"]
            # ojs_book = load_workbook(ojsfile, read_only=False, keep_vba=True)
            # ws = ojs_book['AwardList']
            table: Table = ws.tables["AwardList"]
            table_range: str = table.ref # should return a string A1:B2
            start_cell = table_range.split(':')[0] # should return 'A1'
            coords = coordinate_from_string(start_cell)
            start_col = column_index_from_string(coords[0])
            # start_index = divassignees.index[0]
            table.ref = table.ref[:-1] + str(len(divawards) + 1 + judge_awards)
            for i, row in divawards.iterrows():
                award = row['Award'].replace("Champ", "Champions ").replace("RD", "Robot Design ").replace("CV", "Core Values ").replace("RG", "Robot Game " ).replace("IP", "Innovation Project ")
                award = award.replace("1", "1st Place").replace("2", "2nd Place").replace("3", "3rd Place")
                ws.cell(row=i + 2, column=start_col).value = award
            next_row = i + 1
            for i in range(judge_awards):
                ws.cell(row=i + 2 + next_row, column=start_col).value = "Judges Award " + str(i + 1)

            ojs_book.save(ojsfile)

def set_up_meta_worksheet(tournament:pd.Series, yr: int, seasonName: str):
    for d in ["D1", "D2"]:
        print(f"Setting up meta worksheet for {tournament["Short Name"]} {d}")
        print(f'Season year: {yr}; season name: {seasonName}')
        print(tournament)
        if tournament[d + "_OJS"] is not None:
            print(f'ojs file is not None. Here it is:{tournament[d + "_OJS"]}.')
            ojsfile = dir_path + "\\tournaments\\" + tournament["Short Name"] + "\\" + tournament[d + "_OJS"]
            print(f'Loading ojs workbook {ojsfile}')
            ojs_book = load_workbook(ojsfile, read_only=False, keep_vba=True)
            scriptfile = str(yr) + "-" + seasonName + "-" + tournament["Short Name"] + "-script.html"
            divawards: pd.DataFrame = dfAwards[(dfAwards["Tournament"] == tournament["Short Name"]) & (dfAwards["Div"] == d)]
            divawards = divawards.transpose()
            divawards = divawards.reset_index()
            divawards.columns = ['Award', 'Count']
            ws = ojs_book['Meta']
            # Tournament Year
            ws.cell(row=2, column=2).value = yr
            # FLL Season Title
            ws.cell(row=3, column=2).value = seasonName
            # Tournament Short Name
            ws.cell(row=4, column=2).value = tournament["Short Name"]
            # Tournament Short Name
            ws.cell(row=5, column=2).value = tournament["Long Name"]
            # Completed Script File
            ws.cell(row=6, column=2).value = scriptfile
            # Division
            ws.cell(row=7, column=2).value = 1 if d == "D1" else 2
            # Robot Game
            ws.cell(row=8, column=2).value = divawards.loc[(divawards["Award"] == "RG1") | (divawards["Award"] == "RG2") | (divawards["Award"] == "RG3"), "Count"].sum()
            # Innovation Project
            ws.cell(row=9, column=2).value = divawards.loc[(divawards["Award"] == "IP1") | (divawards["Award"] == "IP2") | (divawards["Award"] == "IP3"), "Count"].sum()
            # Core Values
            ws.cell(row=10, column=2).value = divawards.loc[(divawards["Award"] == "CV1") | (divawards["Award"] == "CV2") | (divawards["Award"] == "CV3"), "Count"].sum()
            # Robot Design
            ws.cell(row=11, column=2).value = divawards.loc[(divawards["Award"] == "RD1") | (divawards["Award"] == "RD2") | (divawards["Award"] == "RD3"), "Count"].sum()
            # Champions
            ws.cell(row=12, column=2).value = divawards.loc[(divawards["Award"] == "Champ1") | (divawards["Award"] == "Champ2") | (divawards["Award"] == "Champ3"), "Count"].sum()
            # Advancing
            ws.cell(row=13, column=2).value = divawards.loc[divawards["Award"] == "ADV", "Count"].values[0]
            # Judges
            ws.cell(row=14, column=2).value = tournament["Judges_1"] + tournament["Judges_2"] + tournament["Judges_3"] + tournament["Judges_4"] + tournament["Judges_5"] + tournament["Judges_6"]
            ojs_book.save(ojsfile)


# # # # # # # # # # # # # # # # # # # # #

print("Checking to make sure files and folders are set up correctly")
extrafilelist: list[str] = [dir_path + "\\script_maker.exe", dir_path + "\\script_template.html.jinja"]
ojsfilelist: list[str] = []
for filename in extrafilelist:
    try:
        if os.path.exists(filename):
            print(f'{filename}... CHECK!',
            tag="info",
            tag_color="green",
            color="green",
        )
    except Exception as e:
        print(
            f'Got an error checking for {filename}\n{e}',
            tag="error",
            tag_color="red",
            color="red",
        )
        input("Press enter to quit...")
        sys.exit(1)

try:
    print(f"Getting tournaments from {dir_path}")
    book = load_workbook(tournament_file, data_only=True)
    ws = book['Tournaments']
    columns, data = read_excel_table(ws, 'TournamentList')
    dfTournaments = pd.DataFrame(data=data, columns=columns)
    seasonYear: int = ws["B1"].value
    seasonName: str = ws["B2"].value
    print(seasonYear, seasonName)

    print("Getting awards")
    ws = book['Awards']
    columns, data = read_excel_table(ws, 'AwardList')
    dfAwards = pd.DataFrame(data=data, columns=columns)
    print(dfAwards)

    print("Getting assignments")
    ws = book['Assignments']
    columns, data = read_excel_table(ws, 'Assignments')
    dfAssignments = pd.DataFrame(data=data, columns=columns)
except Exception as e:
    print(
        f'Could not open the tournament file: {tournament_file}. Check to make sure it is not open in Excel.\n{e}',
        tag="error",
        tag_color="red",
        color="red",
    )
    input("Press enter to quit...")
    sys.exit(1)

for index, row in dfTournaments.iterrows():
    newpath = dir_path + "\\tournaments\\" + row['Short Name']
    judge_award_count = row['Judges_1'] + row['Judges_2'] + row['Judges_3'] + row['Judges_4'] + row['Judges_5'] + row['Judges_6']
    create_folder(newpath)
    copy_files(row)
    set_up_tapi_worksheet(row)
    set_up_award_worksheet(row, judge_award_count)
    set_up_meta_worksheet(row, seasonYear, seasonName)
   
input("Press enter to quit...")
sys.exit(1)

