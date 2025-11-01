# Fatal errors and warnings

Doing any of these will cause a fatal error or warning as described

## File / template / folder errors

* Remove file_list.txt or delete one of the files listed in it -> Fatal: "file_list.txt missing" or "The <file> file is missing..."
* Remove script_template.html.jinja or corrupt it -> Fatal: "Could not read the template file..."
* Leave any ~*.xlsm temp file in the folder (Excel open) -> Fatal: temporary file found message
* Put 0 or >2 OJS files (no div.xlsm or 3+ matches) -> Fatal: "There must be one or two OJS files..."
* Rename an OJS file so it doesn't match expected regex -> Fatal: "Unexpected OJS filename format: ..."

## Meta / table existence errors

* Remove the expected table (Meta, TournamentData, etc.) or rename the worksheet/table -> Exceptions/fatal when read_excel_table fails / when code expects keys.
* Remove keys from Meta (e.g. "Judges", "Advancing", "Completed Script File", "Tournament Long Name") or make them non-numeric where numeric expected -> Missing/invalid values (some are handled as 0, others may cause later fatal/warnings or template render errors).

## Robot game / numeric range checks

* Put an empty cell in Robot Game Scores range C2:E{N} -> Warning/fatal from check_range_for_empty_cells depending on flow (currently fatal branch in run_validations).
* Put non-numeric text in Robot Game Scores -> Fatal/warning from check_range_for_valid_numbers ("not a number")
* Put a number outside 0–545 in Robot Game Scores (e.g. 500 or -5) -> Invalid value message (fatal in validations)

## Judged-input / allowed-values checks

* Leave empty cells in Robot Design Input D2:M{N}, Core Values N2:P, Innovation Project D2:M -> Warning/fatal (empty-cells check)
* Put a numeric value not in allowed_values (e.g. 5 in Robot Design where allowed is [0,1,2,3,4]) -> Invalid values message
* Put non-numeric text in those judged cells -> "not a number" invalid message

## Ranking / missing expected data

* Remove Robot Game Rank or Max Robot Game Score rows (missing ranks) or leave ranks incomplete -> Fatal in build_robot_game_html ("Some robot game scores are missing.")
* Remove Team Number or Team Name values for ranked rows -> Errors when code tries to int() team number or render team name (fatal or template problems)

## Award selection / duplication

* Do not select an expected judged award in Results and Rankings (expected count >0 but missing) -> Warning printed that award is missing (non-fatal)
* Assign the same Award string to multiple teams (duplicate award text) -> Fatal: prints duplicate rows and stops

## Advancing / Judges counts

* Mark more teams as Advance? == "Yes" than the Meta Advancing value -> Fatal: "You have selected more advancing ... than allowed"
* Mark fewer advancing teams than allowed -> Warning (warn_continue)
* No Alt advancing or >1 Alt advancing -> Warning
* Select more total Judges-award named rows than Meta Judges value -> Fatal
* Select fewer Judges awards than allowed -> Warning

## Filename / parsing / load errors

* Make the OJS file unreadable (locked/corrupt) -> Fatal when load_workbook fails
* Put unexpected formats in columns headers (e.g. rename "Award" to "Awards") -> Many lookups will fail or be empty; you will see missing/duplicate/fatal messages depending which code path is hit

## Practical tests to inject quickly

* Empty cell: clear a cell in Robot Game Scores C2 -> triggers empty-cell check
* Bad number: put "abc" in Robot Game Scores C2 -> "not a number"
* Out-of-range: put 999 in Robot Game Scores -> invalid range message
* Disallowed judged value: put 5 in Robot Design input -> invalid-values
* Duplicate award: set Award column value "Champions 1st Place" for two different rows -> duplicate_award fatal
* Too many advancing: set Meta Advancing = 1, but mark 2 teams Advance? = "Yes" -> fatal
* Missing template file or missing file_list.txt -> fatal at startup