"""write_json5.py

Create a sample JSON5 file from a Python dictionary. Includes error handling and
fallback to standard JSON if the optional `json5` package is not installed.

Usage:
    python write_json5.py

Dependencies (optional):
    pip install json5
"""

from pathlib import Path
import json
import sys

def build_sample() -> dict:
    """Return a sample dictionary with nested keys and arrays."""
    return {
  "TOURNAMENT": {
    "notes": "If you have either div1 or div2 set to true, you should set divisions_enabled to true. If your tournament doesn't use divisions at all, set all of these to false. Even if the tournament is a single division (i.e., in a region that uses divisions), set division_enabled to true and either div1 or div2 to true as needed. Setting two_emcees to true will use an alternating color scheme for the closing ceremony script to make it easier for the emcees to read",
    "division_enabled": True,
    "div1": True,
    "div2": True,
    "two_emcees": False,
    "tournament_name": "Sample",
    "season_name": "Unearthed",
    "tournament_date": "2025-12-13",
    "other_required_files": ["script_template.html.jinja"]
  },
  "D1": {
    "notes": "If your tournament uses divisions, set up the awards here. If you set div1 to false above, this section will be ignored. Same for div2 and the D2 section below.",
    "ojs": "2025-vadc-fll-challenge-sample-ojs-norfolk-div1.xlsm",
    "Robot Game": {
      "qty": "2",
      "places": [
        {
          "place": "1",
          "label": "Robot Game 1st Place"
        },
        {
          "place": "2",
          "label": "Robot Game 2nd Place"
        }
      ]
    },
    "Champions": {
      "qty": "2",
      "places": [
        {
          "place": "1",
          "label": "Champions 1st Place"
        },
        {
          "place": "2",
          "label": "Champions 2nd Place"
        }
      ]
    },
    "Innovation Project": {
      "qty": "1",
      "places": [
        {
          "place": "1",
          "label": "Innovation Project 1st Place"
        }
      ]
    },
    "Robot Design": {
      "qty": "1",
      "places": [
        {
          "place": "1",
          "label": "Robot Design 1st Place"
        }
      ]
    },
    "Core Values": {
      "qty": "1",
      "places": [
        {
          "place": "1",
          "label": "Core Values 1st Place"
        }
      ]
    }
  },
  "D2": {
    "ojs2": "2025-vadc-fll-challenge-sample-ojs-norfolk-div2.xlsm",
    "Robot Game": {
      "qty": "2",
      "places": [
        {
          "place": "1",
          "label": "Robot Game 1st Place"
        },
        {
          "place": "2",
          "label": "Robot Game 2nd Place"
        }
      ]
    },
    "Champions": {
      "qty": "2",
      "places": [
        {
          "place": "1",
          "label": "Champions 1st Place"
        },
        {
          "place": "2",
          "label": "Champions 2nd Place"
        }
      ]
    },
    "Innovation Project": {
      "qty": "1",
      "places": [
        {
          "place": "1",
          "label": "Innovation Project 1st Place"
        }
      ]
    },
    "Robot Design": {
      "qty": "1",
      "places": [
        {
          "place": "1",
          "label": "Robot Design 1st Place"
        }
      ]
    },
    "Core Values": {
      "qty": "1",
      "places": [
        {
          "place": "1",
          "label": "Core Values 1st Place"
        }
      ]
    }
  },
  "TOURNAMENT_AWARDS": {
    "notes": "Awards that are presented at the tournament level. The awards below do not break out within divisions. If your tournament does not use divisions, set the ojs value to empty string. ojs entry when used should look like tournament_ojs.xlsm",
    "ojs": "",
    "Robot Game": {
      "qty": "0",
      "places": []
    },
    "Champions": {
      "qty": "0",
      "places": []
    },
    "Innovation Project": {
      "qty": "0",
      "places": []
    },
    "Robot Design": {
      "qty": "0",
      "places": []
    },
    "Core Values": {
      "qty": "0",
      "places": []
    },
    "Judges": {
      "notes": "The judges award is normally presented at the tournament level, even in regions that have divisions",
      "qty": "2",
      "places": [
        {
          "place": "1",
          "label": "Judges 1"
        },
        {
          "place": "2",
          "label": "Judges 2"
        }
      ]
    },
    "Breakthrough Award": {
      "notes": "Other non-core awards. Other awards can be added here, or the ones here can be edited, but that is not recommended. It is better to make sure the tournament spreadsheet has all of the correct awards and this file will automatically generate.",
      "qty": "0",
      "places": [
        {
          "place": "0",
          "label": "Breakthrough Award"
        }
      ]
    },
    "Rising Allstar": {
      "qty": "0",
      "places": [
        {
          "place": "0",
          "label": "Rising Allstar"
        }
      ]
    },
    "Community Leader": {
      "qty": "0",
      "places": [
        {
          "place": "0",
          "label": "Community Leader"
        }
      ]
    }
  }
}

def write_json5(path: Path, data: dict) -> None:
    """Write `data` to `path` as JSON5 when possible, otherwise fallback to JSON.

    The function catches and reports IO/serialization errors.
    """
    try:
        with path.open('w', encoding='utf-8') as fh:
            # json5.dump supports pretty printing via indent
            json.dump(data, fh, indent=2)
        print(f"Wrote JSON5 to {path}")
    except (OSError, IOError) as e:
        print(f"I/O error while writing {path}: {e}", file=sys.stderr)
    except Exception as e:
        print(f"Unexpected error while serializing data: {e}", file=sys.stderr)


def main():
    out = Path('sample_config.json')
    data = build_sample()
    write_json5(out, data)


if __name__ == '__main__':
    main()
