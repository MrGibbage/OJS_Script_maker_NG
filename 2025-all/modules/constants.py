"""Constants used throughout the OJS tournament folder builder."""

# Configuration
SHEET_PASSWORD: str = "skip"
CONFIG_FILENAME: str = "season.json"

# Column names
COL_TEAM_NUMBER: str = "Team #"  # Revert back to "Team #"
COL_TEAM_NAME: str = "Team Name"
COL_COACH_NAME: str = "Coach Name"
COL_POD_NUMBER: str = "Pod Number"
COL_SHORT_NAME: str = "Short Name"
COL_LONG_NAME: str = "Long Name"
COL_OJS_FILENAME: str = "OJS_FileName"
COL_DIVISION: str = "Div"
COL_ADVANCING: str = "ADV"
COL_DATE: str = "Date"
COL_COLUMN_NAME: str = "ColumnName"
COL_DIV_AWARD: str = "DivAward"
COL_SCRIPT_TAG_D1: str = "ScriptTagD1"
COL_SCRIPT_TAG_D2: str = "ScriptTagD2"
COL_SCRIPT_TAG_NODIV: str = "ScriptTagNoDiv"

# Sheet names
SHEET_TEAM_INFO: str = "Team and Program Information"
SHEET_AWARD_DROPDOWNS: str = "AwardListDropdowns"
SHEET_META: str = "Meta"
SHEET_AWARD_DEF: str = "AwardDef"
SHEET_ROBOT_GAME: str = "Robot Game Scores"
SHEET_INNOVATION: str = "Innovation Project Input"
SHEET_ROBOT_DESIGN: str = "Robot Design Input"
SHEET_CORE_VALUES: str = "Core Values Input"
SHEET_RESULTS: str = "Results and Rankings"

# Table names
TABLE_TEAM_LIST: str = "OfficialTeamList"
TABLE_ROBOT_GAME_AWARDS: str = "RobotGameAwards"
TABLE_AWARD_DROPDOWNS: str = "AwardListDropdowns"
TABLE_META: str = "Meta"
TABLE_AWARD_DEF: str = "AwardDef"
TABLE_ROBOT_GAME: str = "RobotGameScores"
TABLE_INNOVATION: str = "InnovationProjectResults"
TABLE_ROBOT_DESIGN: str = "RobotDesignResults"
TABLE_CORE_VALUES: str = "CoreValuesResults"
TABLE_TOURNAMENT_DATA: str = "TournamentData"

# File structure
FILE_CLOSING_CEREMONY: str = "closing_ceremony.html"

# Award columns
AWARD_COLUMN_PREFIX_JUDGED: str = "J_"
AWARD_COLUMN_ROBOT_GAME: str = "P_AWD_RG"
AWARD_LABEL_PREFIX: str = "Label"

# Required columns for each table (for validation)
# Maps table name to list of required column names
REQUIRED_COLUMNS = {
    "OfficialTeamList": ["Team #", "Team Name", "Coach Name"],
    "RobotGameAwards": ["Robot Game Awards"],
    "AwardListDropdowns": ["Award"],
    "Meta": ["Key", "Value"],
    "AwardDef": ["ColumnName"],  # Label columns are dynamic
    "RobotGameScores": ["Team #"],
    "InnovationProjectResults": ["Team #"],
    "RobotDesignResults": ["Team #"],
    "CoreValuesResults": ["Team #"],
    "TournamentData": ["Team #"],  # Changed from "Team Number"
}
