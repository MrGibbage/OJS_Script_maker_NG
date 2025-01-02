# pip install pandas
import pandas as pd

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, get_column_letter, coordinate_from_string
from openpyxl.worksheet.table import Table

# pip install pyyaml
import yaml
import warnings

import os, sys, re, glob

# pip install Jinja2
from jinja2 import Environment, FileSystemLoader, select_autoescape

from typing import List, Dict

# To create windows exe executable, run
# .venv\Scripts\pyinstaller.exe -F script_maker.py
# in the project folder. The executable will be saved in the 'dist'
# folder. Just copy it up to the project folder.
# Double-click to run.

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

print("Building Closing Ceremony script")

templateLoader = FileSystemLoader(searchpath="2024/")
templateEnv = Environment(loader=templateLoader)
TEMPLATE_FILE = "script_template.html.jinja"
template = templateEnv.get_template(TEMPLATE_FILE)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

ordinals = ["1st", "2nd", "3rd", "4th", "5th"]
dir_path = os.path.dirname(os.path.realpath(__file__))

# print("Opening yaml data file: " + yaml_data_file)
# with open(yaml_data_file) as f:
#     dict = yaml.load(f, Loader=yaml.FullLoader)

awards = {
    "Robot Game": "",
    "Champions": "",
    "Innovation Project": "",
    "Robot Design": "",
    "Core Values": "",
    "Judges": "",
}
rg_html = {}
dataframes = {}
# divAwards key will be 1 and/or 2, and the value will be the awardHtml dictionaries
# yes, it is a dictionary of dictionaries
divAwards: Dict[int, Dict[str, str]] = {}
divAwards[1] = {}
divAwards[2] = {} 
# awardHtml key will be the award name, and the values will be the html
awardHtml = {}
advancingDf = {}
advancingHtml = {}
awardCounts = {}
judgesAwardDf = {}
judgesAwardHtml = {}
judgesAwardTotalCount = 0

divisions = [1, 2]
for award in awards:
    for div in divisions:
        divAwards[div][award] = ''

removing = []
if len(glob.glob(f"{dir_path}\\~*.xlsm")) > 0:
    print("Found temporary file(s) indicating that you have one or more spreadsheets open in Excel. Please close Excel and retry.")
    input("Press enter to quit...")
    sys.exit(1)
directory_list: list[str] = glob.glob(f"{dir_path}\\*div*.xlsm")
print(directory_list)

for award in awards:
    awardCounts[award] = 0

for tourn_filename in directory_list:
    regex: str = r"([0-9]{4}-vadc-fll-challenge-.*)(-ojs-)(.*)-(div[1,2])(.xlsm)$"
    # checking the file name for the division. The filename includes the path
    # so remove the path by using the length of dir_path
    m0 = re.search(regex, tourn_filename[len(dir_path) + 1:])
    div = int(m0.group(4)[-1])
    print(f"Division {div}")

    book = load_workbook(tourn_filename, data_only=True)
    ws = book['Meta']
    columns, data = read_excel_table(ws, 'Meta')
    dfMeta = pd.DataFrame(data=data, columns=["Key", "Value"])
    for award in awards:
        awardCounts[award] = dfMeta.loc[dfMeta['Key'] == award, 'Value'].values[0]
    print(awardCounts)
    ws = book["Results and Rankings"]
    columns, data = read_excel_table(ws, 'TournamentData')
    dfRankings = pd.DataFrame(data=data, columns=columns)
    print(dfRankings)

    # Robot Game
    rg_html[div] = ""
    for i in reversed(range(awardCounts["Robot Game"])):
        teamNum = dfRankings.loc[dfRankings['Robot Game Rank'] == i + 1, 'Team Number'].values[0]
        teamName = dfRankings.loc[dfRankings['Robot Game Rank'] == i + 1, 'Team Name'].values[0]
        score = int(dfRankings.loc[dfRankings['Robot Game Rank'] == i + 1, 'Max Robot Game Score'].values[0])
        rg_html[div] += (
            "<p>With a score of "
            + str(score)
            + " points, the Division " + str(div) + " "
            + ordinals[i]
            + " place award goes to team number "
            + str(int(teamNum))
            + ", "
            + teamName
            + "</p>\n"
        )
    
    # Judged awards
    for award in awards:
        print(award)
        if award == "Robot Game" or award == "Judges":
            continue
        for i in reversed(range(awardCounts[award])):
            thisAward = award + " " + ordinals[i] + " Place"
            print(f'Division {div}, {thisAward}')
            print(dfRankings.loc[dfRankings['Award'] == thisAward, 'Team Number'])
            teamNum = dfRankings.loc[dfRankings['Award'] == thisAward, 'Team Number'].values[0]
            teamName = dfRankings.loc[dfRankings['Award'] == thisAward, 'Team Name'].values[0]
            thisText = (
                "<p>The Division " + str(div) + " "
                + ordinals[i]
                + " place " + award + " award goes to team number "
                + str(int(teamNum))
                + ", "
                + teamName
                + "</p>\n"
            )
            divAwards[int(div)][award] += thisText

    advancingDf[int(div)] = dfRankings[dfRankings["Advance?"] == "Yes"]
    advancingHtml[div] = ""
    # print(advancingDf[div])
    for index, row in advancingDf[int(div)].iterrows():
        try:
            teamNum = str(int(row['Team Number']))
        except:
            teamNum = ""
        teamName = row["Team Name"]
        advancingHtml[div] += "<p>(Div " + str(div) + ") Team number " + teamNum + ", " + teamName + "</p>\n"
    print(f'Advancing: {advancingHtml[div]}')

    print(dfRankings)
    # Judges Awards
    allowedJudgesAwardCount = dfMeta.loc[dfMeta['Key'] == "Judges", 'Value'].values[0]
    judgesAwardDf[int(div)] = dfRankings[(dfRankings["Award"].str.startswith("Judges", na=False))]
    print(judgesAwardDf)
    judgesAwardHtml[div] = ""
    # print(advancingDf[div])
    for index, row in judgesAwardDf[int(div)].iterrows():
        try:
            teamNum = str(int(row['Team Number']))
        except:
            teamNum = ""
        teamName = row["Team Name"]
        judgesAwardHtml[div] += "<p>(Div " + str(div) + ") Team number " + teamNum + ", " + teamName + "</p>\n"
        judgesAwardTotalCount += 1
    
print(f'Total number of judges awards: Div 1 = {len(judgesAwardDf[1])}; Div 2 = {len(judgesAwardDf[2])}; total = {judgesAwardTotalCount}')
if judgesAwardTotalCount > allowedJudgesAwardCount:
    print(f'You have selected too many judges awards. You are permitted a total of {allowedJudgesAwardCount} judges awards. If your tournament has two divisions, that total number is across both divisions. In other words, if you are allowed three judges awards, you could have three from Div 1 or two from Div 1 and one from Div 2, etc.')
    input("Press enter to quit...")
    sys.exit(1)


print("Rendering the script")
out_text = template.render(
    tournament_name = dfMeta.loc[dfMeta['Key'] == "Tournament Long Name", 'Value'].values[0], 
    rg_div1_list = rg_html[1],
    rg_div2_list = rg_html[2],
    rd_div1_list = divAwards[1]["Robot Design"],
    rd_div2_list = divAwards[2]["Robot Design"],
    rd_this_them = "This team" if awardCounts["Robot Design"] == 1 else "These teams",
    ip_div1_list = divAwards[1]["Innovation Project"],
    ip_div2_list = divAwards[2]["Innovation Project"],
    ip_this_them = "This team" if awardCounts["Innovation Project"] == 1 else "These teams",
    cv_div1_list = divAwards[1]["Core Values"],
    cv_div2_list = divAwards[2]["Core Values"],
    cv_this_them = "This team" if awardCounts["Core Values"] == 1 else "These teams",
    ja_count = judgesAwardTotalCount,
    ja_list = judgesAwardHtml[1] + judgesAwardHtml[2],
    ja_go_goes = "The Judges Awards go to teams" if judgesAwardTotalCount > 1 else "The Judges Award goes to team",
    champ_div1_list = divAwards[1]["Champions"],
    champ_div2_list = divAwards[2]["Champions"],
    adv_div1_list = advancingHtml[1],
    adv_div2_list = advancingHtml[2],
)

# print(out_text)
with open(dfMeta.loc[dfMeta['Key'] == "Completed Script File", 'Value'].values[0], "w") as fh:
    fh.write(out_text)

print("All done! The script hase been saved as " + dfMeta.loc[dfMeta['Key'] == "Completed Script File", 'Value'].values[0] + ".")
print("It is saved in the same folder with the other OJS files.")
print("Double-click the script file to view it on this computer,")
print("or email it to yourself and view it on a phone or tablet.")
input("Press enter to quit...")