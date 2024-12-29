# This program will run a pre-check on an OJS folder to make
# sure all of the files are set up correctly
# These are the checks the program will validate
import os, sys, re

# pip install pandas
import pandas as pd

# pip install print-color
from print_color import print

print("Getting a directory listing", tag="info", tag_color="white", color="white")
try:
    directory_list: list[str] = os.listdir()
    print(
        f"Found {len(directory_list)} files in the folder.",
        tag="info",
        tag_color="white",
        color="white",
    )
except Exception as e:
    print(
        f"*-*-* Could not get a directory list. We got this error: {e}",
        tag="error",
        tag_color="red",
        color="red",
    )
    input("Press enter to quit...")
    sys.exit(1)

print("Looking for the OJS spreadsheets", tag="info", tag_color="white", color="white")
xlsm_files: list[str] = [s for s in directory_list if s.endswith(".xlsm")]

if len(xlsm_files) > 0:
    print(
        f"Found these spreadsheets: {xlsm_files}",
        tag="info",
        tag_color="white",
        color="white",
    )
else:
    print(
        "*-*-* Did not find any spreadsheets. Be sure to run this program from the same folder where the spreadsheets are saved.",
        tag="error",
        tag_color="red",
        color="red",
    )
    input("Press enter to quit...")
    sys.exit(1)

if len(xlsm_files) > 2:
    print(
        f"*-*-* There should be only one or two spreadsheets in the folder. We found {len(xlsm_files)} spreadsheets.",
        tag="error",
        tag_color="red",
        color="red",
    )
    print(
        f"*-*-* Perhaps you have one or more of the spreadsheets open, "
        "which will add temprary files with .xlsm extensions.",
        tag="error",
        tag_color="red",
        color="red",
    )
    input("Press enter to quit...")
    sys.exit(1)

regex: str = "^([0-9]{4}-vadc-fll-challenge-.*)(-ojs-)(.*)-(div[1,2])(\.xlsm)$"

for f in xlsm_files:
    print(f"Checking {f}", tag="info", tag_color="white", color="white")
    m = re.search(regex, f)
    try:
        print(m.groups(), tag="info", tag_color="white", color="white")
        if not ((m.group(2)) == "-ojs-" and m.group(4).startswith("div")):
            print(
                f"*-*-* {f} is not named correctly",
                tag="error",
                tag_color="red",
                color="red",
            )
            input("Press enter to quit...")
            sys.exit(1)
    except Exception as e:
        print(
            f"*-*-* {f} is not named correctly",
            tag="error",
            tag_color="red",
            color="red",
        )
        if (f[:1]=="~"):
            print(
                f"It looks like {f} is a temporary file, which suggests you may have an OJS file open in Excel",
                tag="error",
                tag_color="red",
                color="red",
            )
        print(
            "OJS files should be named with this pattern (all lowercase, "
            "no special characters or spaces)"
        )
        print("year-vadc-fll-challenge-season_name-ojs-tournament_name-div#.xlsm")
        print("For example, 2024-vadc-fll-challenge-submerged-ojs-norview-div1.xlsm")
        print("Where vadc-fll-challenge is always the same")
        input("Press enter to quit...")
        sys.exit(1)
divlist: list[str] = ["div1", "div2"]

div: list[str] = []
base_file_name: str = ""
if len(xlsm_files) == 2:
    print(
        "Checking to see if there is a div1 and a div2",
        tag="info",
        tag_color="white",
        color="white",
    )
    m0 = re.search(regex, xlsm_files[0])
    m1 = re.search(regex, xlsm_files[1])
    div.append(m0.group(4))
    div.append(m1.group(4))
    print("Found", div[0], div[1], tag="info", tag_color="white", color="white")
    if "div1" in div and "div2" in div:
        print(
            "Good. There are two divisions: div1 and div2",
            tag="info",
            tag_color="white",
            color="white",
        )
    else:
        print(
            "*-*-* There should be two different divisions",
            tag="error",
            tag_color="red",
            color="red",
        )
        input("Press enter to quit...")
        sys.exit(1)
    if m0.group(1) != m1.group(1):
        print(
            f"*-*-* {m0.group(1)} does not match {m1.group(1)}",
            tag="error",
            tag_color="red",
            color="red",
        )
        input("Press enter to quit...")
        sys.exit(1)
    if m0.group(3) != m1.group(3):
        print(
            f"*-*-* {m0.group(3)} does not match {m1.group(3)}",
            tag="error",
            tag_color="red",
            color="red",
        )
        input("Press enter to quit...")
        sys.exit(1)
    if m0.group(5) != m1.group(5):
        print(
            f"*-*-* {m0.group(5)} does not match {m1.group(5)}",
            tag="error",
            tag_color="red",
            color="red",
        )
        input("Press enter to quit...")
        sys.exit(1)
    base_file_name = m0.group(1) + m0.group(2) + m0.group(3) + "-"
    print(
        "Files appear to be named correctly.",
        tag="success",
        tag_color="green",
        color="white",
    )

print(f"Base file name: {base_file_name}", tag="info", tag_color="white", color="white")

if len(xlsm_files) == 1:
    print(
        "Checking to see if it is a div1 or div2",
        tag="info",
        tag_color="white",
        color="white",
    )
    m = re.search(regex, xlsm_files[0])
    div = m.group(4)
    print("Found", div, tag="info", tag_color="white", color="white")
    if div in divlist:
        print(
            f"Good. Found {m.group(4)}", tag="success", tag_color="green", color="white"
        )
    else:
        print(
            f"*-*-* Neither div1 or div2; found {div}",
            tag="error",
            tag_color="red",
            color="red",
        )
        input("Press enter to quit...")
        sys.exit(1)
    divlist.remove("div1" if div == "div2" else "div2")
    print(
        "File appears to be named correctly.",
        tag="success",
        tag_color="green",
        color="white",
    )

print(divlist, tag="info", tag_color="white", color="white")

# Check to see if any of the spreadsheets are open in Excel
for division in divlist:
    this_ojs_filename = base_file_name + division + ".xlsm"
    print(
        f"Checking {this_ojs_filename} to see if it is open",
        tag="info",
        tag_color="white",
        color="white",
    )
    try:
        # https://stackoverflow.com/questions/6825994/check-if-a-file-is-open-in-python
        os.rename(this_ojs_filename, this_ojs_filename)
        print(
            f"{this_ojs_filename} is correctly closed.",
            tag="info",
            tag_color="white",
            color="white",
        )
    except:
        print(
            f"{this_ojs_filename} seems to be open. Do you have this file open in Excel?",
            tag="error",
            tag_color="red",
            color="red",
        )
        sys.exit(1)

dataframes = {}

try:
    for division in divlist:
        this_ojs_filename = base_file_name + division + ".xlsm"
        print(
            f"Opening {this_ojs_filename} to read Results and Rankings",
            tag="info",
            tag_color="white",
            color="white",
        )
        dataframes[division] = pd.read_excel(
            this_ojs_filename,
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
        print(
            "There should be no errors or warnings. All rows below should "
            "have team data.",
            tag="info",
            tag_color="white",
            color="white",
        )
        print(
            "Robot game scores should be all 0",
            tag="info",
            tag_color="white",
            color="white",
        )
        print(
            "Robot game ranks should be all 1",
            tag="info",
            tag_color="white",
            color="white",
        )
        print(
            "Award and Advance? should be all NaN",
            tag="info",
            tag_color="white",
            color="white",
        )
        print(
            "All team numbers should be integers and there should not be "
            "any team names NaN",
            tag="info",
            tag_color="white",
            color="white",
        )
        print(dataframes[division])
except Exception as e:
    print(
        f"There was an error reading the OJS spreadsheet {this_ojs_filename}",
        tag="error",
        tag_color="red",
        color="red",
    )
    print(f"The error message was {e}", tag="error", tag_color="red", color="red")
    input("Press enter to quit...")
    sys.exit(1)

print(
    "Checking for null values in the team numbers",
    tag="info",
    tag_color="white",
    color="white",
)
try:
    for division in divlist:
        if dataframes[division]["Team Number"].isnull().values.any():
            print(
                f"Found a null in the {division} OJS Team Numbers",
                tag="error",
                tag_color="red",
                color="red",
            )
            sys.exit(1)
except Exception as e:
    print(f"There was an error {e}", tag="error", tag_color="red", color="red")
    input("Press enter to quit...")
    sys.exit(1)

print(
    "Checking for null values in the team names",
    tag="info",
    tag_color="white",
    color="white",
)
try:
    for division in divlist:
        if dataframes[division]["Team Name"].isnull().values.any():
            print(
                f"Found a null in the {division} OJS Team Names",
                tag="error",
                tag_color="red",
                color="red",
            )
            input("Press enter to quit...")
            sys.exit(1)
except Exception as e:
    print(f"There was an error {e}", tag="error", tag_color="red", color="red")
    input("Press enter to quit...")
    sys.exit(1)

print(
    "Checking for zeroes in the Robot Game Scores",
    tag="info",
    tag_color="white",
    color="white",
)
try:
    for division in divlist:
        if (dataframes[division]["Max Robot Game Score"] != 0).all():
            print(
                f"Found a non-zero in the {division} OJS Max Scores",
                tag="error",
                tag_color="red",
                color="red",
            )
            input("Press enter to quit...")
            sys.exit(1)
except Exception as e:
    print(f"There was an error {e}", tag="error", tag_color="red", color="red")
    input("Press enter to quit...")
    sys.exit(1)

print(
    "Checking for ones in the Robot Game Rank",
    tag="info",
    tag_color="white",
    color="white",
)
try:
    for division in divlist:
        if (dataframes[division]["Robot Game Rank"] != 1).all():
            print(
                f"Found a non-zero in the {division} OJS Robot Game Ranks",
                tag="error",
                tag_color="red",
                color="red",
            )
            input("Press enter to quit...")
            sys.exit(1)
except Exception as e:
    print(f"There was an error {e}", tag="error", tag_color="red", color="red")
    input("Press enter to quit...")
    sys.exit(1)

print(
    "Checking for NaNs in the Award column",
    tag="info",
    tag_color="white",
    color="white",
)
try:
    for division in divlist:
        if not(dataframes[division]["Award"].isnull().values.all()):
            print(
                f"Found a non-NaN in the {division} OJS Awards",
                tag="error",
                tag_color="red",
                color="red",
            )
            input("Press enter to quit...")
            sys.exit(1)
except Exception as e:
    print(f"There was an error {e}", tag="error", tag_color="red", color="red")
    input("Press enter to quit...")
    sys.exit(1)

print(
    "Checking for NaNs in the Advance? column",
    tag="info",
    tag_color="white",
    color="white",
)
try:
    for division in divlist:
        if not(dataframes[division]["Advance?"].isnull().values.all()):
            print(
                f"Found a non-NaN in the {division} OJS Advance column",
                tag="error",
                tag_color="red",
                color="red",
            )
            input("Press enter to quit...")
            sys.exit(1)
except Exception as e:
    print(f"There was an error {e}", tag="error", tag_color="red", color="red")
    input("Press enter to quit...")
    sys.exit(1)


#
# TODO
# check columns in tables/tabs
# check if all awards are set up correctly
# check if sheets are protected with a password
# check meta information
