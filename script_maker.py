# pip install pandas
import pandas as pd

# pip install pyyaml
import yaml
import warnings

# pip install Jinja2
from jinja2 import Environment, FileSystemLoader, select_autoescape

# To create windows exe executable, run
# .venv\Scripts\pyinstaller.exe -F script_maker.py
# in the project folder. The executable will be saved in the 'dist'
# folder. Just copy it up to the project folder.
# Double-click to run.

print("Building Closing Ceremony script")

templateLoader = FileSystemLoader(searchpath="./")
templateEnv = Environment(loader=templateLoader)
TEMPLATE_FILE = "script_template-2MC.html.jinja"
template = templateEnv.get_template(TEMPLATE_FILE)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

yaml_data_file = "meta.yaml"

ordinals = ["1st", "2nd", "3rd", "4th", "5th"]

print("Opening yaml data file: " + yaml_data_file)
with open(yaml_data_file) as f:
    dict = yaml.load(f, Loader=yaml.FullLoader)

awards = {
    "Champions": "",
    "Innovation Project": "",
    "Robot Design": "",
    "Core Values": "",
}
rg_html = {}
dataframes = {}
# divAwards key will be 1 and/or 2, and the value will be the awardHtml dictionaries
# yes, it is a dictionary of dictionaries
divAwards = {}
divAwards[1] = {}
divAwards[2] = {} 
# awardHtml key will be the award name, and the values will be the html
awardHtml = {}
advancingDf = {}
advancingHtml = {}
awardCounts = {}

divisions = [1, 2]
removing = []
for div in divisions:
    if dict["div" + str(div) + "_ojs_file"] is not None:
        dataframes[div] = pd.read_excel(
            dict["div" + str(div) + "_ojs_file"],
            sheet_name="Results and Rankings",
            header=1,
            usecols=[
                "Team Number",
                "Team Name",
                "Max Robot Game Score",
                "Robot Game Rank",
                "Award",
                "Advance?",
            ],
        )
    else:
        # make a list of the elements to remove
        # I can't just remove the element from divisions[] while iterating
        # on it because the loops will end early
        removing.append(div)
        rg_html[div] = ""
        awardHtml[div] = ""
        divAwards[div]["Robot Design"] = ""
        divAwards[div]["Innovation Project"] = ""
        divAwards[div]["Core Values"] = ""
        divAwards[div]["Champions"] = ""
        advancingHtml[div] = ""

for award in awards:
    awardCounts[award] = 0

# now that we are done looping over the divisions, remove any items
# that were marked for deletion
# print(removing)
print("We have these divisions: " + str(divisions))
for item in removing:
    divisions.remove(item)
print("Get the top robot game scores")
# print(divisions)
# Robot Game
for div in divisions:
    rg_html[div] = ""
    print("Getting top scores for division " + str(div))
    for i in reversed(range(int(dict["Division " + str(div) + " Robot Game"]))):
        print(ordinals[i] + " Place")
        teamNum = int(
            dataframes[div].loc[
                dataframes[div]["Robot Game Rank"] == i + 1, "Team Number"
            ].iloc[0]
        )
        teamName = dataframes[div].loc[
            dataframes[div]["Robot Game Rank"] == i + 1, "Team Name"
        ].iloc[0]
        score = int(
            dataframes[div].loc[
                dataframes[div]["Robot Game Rank"] == i + 1, "Max Robot Game Score"
            ].iloc[0]
        )
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

# print(rg_html)
# Judged awards
for div in divisions:
    divAwards[div] = {}
    print("Getting awards for division " + str(div))
    for award in awards.keys():
        print("Getting the " + award + " award")
        divAwards[div][award] = ""
        awardCounts[award] += int(dict["Division " + str(div) + " " + award])
        for i in reversed(range(int(dict["Division " + str(div) + " " + award]))):
            # divAwards[div][award] = ""
            print(award + " " + ordinals[i] + " Place")
            teamNum = int(
                dataframes[div].loc[
                    dataframes[div]["Award"] == award + " " + ordinals[i] + " Place",
                    "Team Number",
                ].iloc[0]
            )
            teamName = dataframes[div].loc[
                dataframes[div]["Award"] == award + " " + ordinals[i] + " Place",
                "Team Name",
            ].iloc[0]
            thisText = (
                "<p>The Division " + str(div) + " "
                + ordinals[i]
                + " place " + award + " award goes to team number "
                + str(int(teamNum))
                + ", "
                + teamName
                + "</p>\n"
            )
            divAwards[div][award] += thisText

print("Getting the advancing teams")
# Advancing
for div in divisions:
    advancingDf[div] = dataframes[div][dataframes[div]["Advance?"] == "Yes"][
        ["Team Number", "Team Name"]
    ]

# print(advancingDf[1])
# print(advancingDf[2])

# # From https://stackoverflow.com/questions/18695605/how-to-convert-a-dataframe-to-a-dictionary
# div1advancingDict = div1advancingDf.set_index("Team Number").to_dict("dict")
# div2advancingDict = div2advancingDf.set_index("Team Number").to_dict("dict")
# # print(div2advancingDict["Team Name"])

for div in divisions:
    advancingHtml[div] = ""
    # print(advancingDf[div])
    for index, row in advancingDf[div].iterrows():
        try:
            teamNum = str(int(row['Team Number']))
        except:
            teamNum = ""
        teamName = row["Team Name"]
        advancingHtml[div] += "<p>(Div " + str(div) + ") Team number " + teamNum + ", " + teamName + "</p>\n"

ja_count = int(dict["Judges Awards"])

print("Rendering the script")
out_text = template.render(
    tournament_name = dict["tournament_name"], 
    volunteer_award_justification = dict["Volunteer Justification"],
    volunteer_awardee_name = dict["Volunteer Awardee"],
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
    ja_count = ja_count,
    ja_list = str(dict["Judges Awardees"]),
    ja_go_goes = "The Judges Awards go to teams" if ja_count > 1 else "The Judges Award goes to team",
    special_guests = str(dict["Special Guests"]),
    champ_div1_list = divAwards[1]["Champions"],
    champ_div2_list = divAwards[2]["Champions"],
    adv_div1_list = advancingHtml[1],
    adv_div2_list = advancingHtml[2],
)

# print(out_text)
with open(dict["complete_script_file"], "w") as fh:
    fh.write(out_text)

print("All done! The script hase been saved as " + dict["complete_script_file"] + ".")
print("It is saved in the same folder with the other OJS files.")
print("Double-click the script file to view it on this computer,")
print("or email it to yourself and view it on a phone or tablet.")
input("Press enter to quit...")