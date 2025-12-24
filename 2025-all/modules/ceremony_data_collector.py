"""Collects award and team data from OJS files for ceremony script generation."""

import logging
import pandas as pd
from typing import Dict, List, Tuple
import os

from .constants import (
    SHEET_RESULTS, TABLE_TOURNAMENT_DATA, SHEET_TEAM_INFO, TABLE_TEAM_LIST,
    COL_TEAM_NUMBER, COL_TEAM_NAME
)
from .excel_operations import read_table_as_df

logger = logging.getLogger("ceremony_generator")


class AwardWinner:
    """Represents a team that won an award."""
    
    def __init__(self, team_number: int, team_name: str, award_name: str, 
                 place: int, division: str = "", score: float = None):
        self.team_number = team_number
        self.team_name = team_name
        self.award_name = award_name
        self.place = place  # 1, 2, 3, etc.
        self.division = division  # "Division 1", "Division 2", or ""
        self.score = score  # For Robot Game awards
    
    def __repr__(self):
        score_str = f", score={self.score}" if self.score is not None else ""
        div_str = f", {self.division}" if self.division else ""
        return f"AwardWinner({self.team_number} {self.team_name}, {self.award_name} {self.place}{div_str}{score_str})"


class CeremonyDataCollector:
    """Collects data from OJS files for ceremony script generation."""
    
    def __init__(self, config: dict):
        self.config = config
        self.using_divisions = config["INFO"]["using_divisions"]
        self.awards = config["AWARDS"]
        self.warnings: List[str] = []
    
    def collect_team_list(self, ojs_path: str, division: str = "") -> List[Tuple[int, str]]:
        """Collect all teams from TournamentData table."""
        logger.info(f"Collecting team list{' for ' + division if division else ''}")
        
        try:
            df = read_table_as_df(ojs_path, SHEET_RESULTS, TABLE_TOURNAMENT_DATA)
            
            teams = []
            for _, row in df.iterrows():
                team_num = int(row[COL_TEAM_NUMBER])
                team_name = str(row[COL_TEAM_NAME])
                teams.append((team_num, team_name))
            
            teams.sort(key=lambda x: x[0])
            logger.debug(f"Collected {len(teams)} teams")
            return teams
        except Exception as e:
            logger.error(f"Error collecting team list: {e}")
            import traceback
            logger.debug(traceback.format_exc())
            return []
    
    def collect_advancing_teams(self, ojs_path: str, division: str = "") -> List[Tuple[int, str]]:
        """Collect teams marked as advancing."""
        logger.info(f"Collecting advancing teams{' for ' + division if division else ''}")
        
        try:
            df = read_table_as_df(ojs_path, SHEET_RESULTS, TABLE_TOURNAMENT_DATA)
            
            advancing = []
            for _, row in df.iterrows():
                advance_status = str(row.get("Advance?", "")).strip()
                if advance_status.upper() == "YES":
                    team_num = int(row[COL_TEAM_NUMBER])
                    team_name = str(row[COL_TEAM_NAME])
                    advancing.append((team_num, team_name))
            
            advancing.sort(key=lambda x: x[0])
            logger.debug(f"Collected {len(advancing)} advancing teams")
            return advancing
        except Exception as e:
            logger.error(f"Error collecting advancing teams: {e}")
            import traceback
            logger.debug(traceback.format_exc())
            return []
    
    def collect_robot_game_awards(self, ojs_path: str, award_count: int, 
                                   division: str = "") -> List[AwardWinner]:
        """Collect Robot Game award winners based on rank."""
        logger.info(f"Collecting {award_count} Robot Game awards{' for ' + division if division else ''}")
        
        try:
            df = read_table_as_df(ojs_path, SHEET_RESULTS, TABLE_TOURNAMENT_DATA)
            
            # Get teams with Robot Game Rank, sorted by rank
            rg_teams = df[df["Robot Game Rank"].notna()].copy()
            rg_teams = rg_teams.sort_values("Robot Game Rank")
            
            winners = []
            for idx, (_, row) in enumerate(rg_teams.head(award_count).iterrows()):
                place = idx + 1
                team_num = int(row[COL_TEAM_NUMBER])
                team_name = str(row[COL_TEAM_NAME])
                score = float(row["Max Robot Game Score"])
                
                winner = AwardWinner(
                    team_number=team_num,
                    team_name=team_name,
                    award_name="Robot Game",
                    place=place,
                    division=division,
                    score=score
                )
                winners.append(winner)
            
            winners.reverse()
            logger.debug(f"Collected {len(winners)} Robot Game winners")
            return winners
        except Exception as e:
            logger.error(f"Error collecting Robot Game awards: {e}")
            import traceback
            logger.debug(traceback.format_exc())
            return []
    
    def collect_judged_awards(self, ojs_path: str, award_def: dict, 
                               award_labels: List[str], division: str = "", 
                               ojs_filename: str = "") -> List[AwardWinner]:
        """Collect judged award winners by matching Award column to labels."""
        award_name = award_def["Name"]
        is_div_award = award_def["DivAwd"]
        logger.info(f"Collecting {award_name} awards{' for ' + division if division else ''}")
        
        try:
            df = read_table_as_df(ojs_path, SHEET_RESULTS, TABLE_TOURNAMENT_DATA)
            
            winners = []
            for place, label in enumerate(award_labels, start=1):
                matches = df[df["Award"] == label]
                
                if len(matches) == 0:
                    if is_div_award:
                        context = f"{division}" if division else f"{ojs_filename}"
                        self.warnings.append(
                            f"{context}: {award_name} award '{label}' not selected"
                        )
                    continue
                
                if len(matches) > 1:
                    context = f"{division}" if division else f"{ojs_filename}"
                    self.warnings.append(
                        f"{context}: {award_name} award '{label}' selected {len(matches)} times (expected 1)"
                    )
                
                for _, row in matches.iterrows():
                    team_num = int(row[COL_TEAM_NUMBER])
                    team_name = str(row[COL_TEAM_NAME])
                    
                    winner = AwardWinner(
                        team_number=team_num,
                        team_name=team_name,
                        award_name=award_name,
                        place=place,
                        division=division
                    )
                    winners.append(winner)
            
            winners.reverse()
            logger.debug(f"Collected {len(winners)} {award_name} winners")
            return winners
        except Exception as e:
            logger.error(f"Error collecting {award_name} awards: {e}")
            import traceback
            logger.debug(traceback.format_exc())
            return []
    
    def _ordinal(self, n: int) -> str:
        """Convert number to ordinal string (1 -> '1st', 2 -> '2nd', etc.)."""
        if 10 <= n % 100 <= 20:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
        return f"{n}{suffix}"
    
    def format_winners_as_html(self, winners: List[AwardWinner], include_score: bool = False) -> str:
        """Format list of winners as HTML paragraphs.
        
        Args:
            winners: List of AwardWinner objects (should be in reverse order already)
            include_score: Whether to include score in output (for Robot Game)
            
        Returns:
            HTML string with <p> tags
        """
        if not winners:
            return ""
        
        html_lines = []
        for winner in winners:
            # Build the output string
            parts = []
            
            # Division (if applicable)
            if winner.division:
                parts.append(f"The {winner.division}")
            else:
                parts.append("The")
            
            # Place
            parts.append(f"{self._ordinal(winner.place)} place")
            
            # Award name
            parts.append(f"{winner.award_name} award")
            
            # Score (for Robot Game)
            if include_score and winner.score is not None:
                parts.append(f"with a score of {int(winner.score)} points")
            
            # Team info
            parts.append(f"goes to team number {winner.team_number}, {winner.team_name}")
            
            html_lines.append(f"<p>{' '.join(parts)}</p>")
        
        return "\n".join(html_lines)
    
    def format_team_list_as_html(self, teams: List[Tuple[int, str]]) -> str:
        """Format team list as HTML paragraphs.
        
        Args:
            teams: List of (team_number, team_name) tuples
            
        Returns:
            HTML string with <p> tags
        """
        if not teams:
            return ""
        
        html_lines = []
        for team_num, team_name in teams:
            html_lines.append(f"<p>Team {team_num}, {team_name}</p>")
        
        return "\n".join(html_lines)
