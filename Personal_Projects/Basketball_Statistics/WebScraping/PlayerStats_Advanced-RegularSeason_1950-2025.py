import time
import random
import requests
import pandas as pd
from io import StringIO
from urllib import robotparser
from bs4 import BeautifulSoup
import os

BASE = "https://www.basketball-reference.com"
ROBOT_URL = BASE + "/robots.txt"
UA = "AdvScraper/9.1 (+https://example.com/info)"
START, END = 1950, 2025
OUT_CSV = f"nba_advanced_{START}_{END}.csv"

# robots.txt
rp = robotparser.RobotFileParser()
rp.set_url(ROBOT_URL)
rp.read()
DELAY = rp.crawl_delay(UA) or 3


def fetch(path):
    url = BASE + path
    if not rp.can_fetch(UA, path):
        return None
    time.sleep(DELAY + random.random())
    r = requests.get(url, headers={"User-Agent": UA})
    return r.text if r.status_code == 200 else None


def parse_season(html, season):
    # 1) Strip comments
    cleaned = html.replace("<!--", "").replace("-->", "")
    # 2) Read tables and find advanced by "PER"
    tables = pd.read_html(StringIO(cleaned))
    adv_df = next((df for df in tables if "PER" in df.columns), None)
    if adv_df is None:
        return pd.DataFrame()
    if "Rk" in adv_df.columns:
        adv_df = adv_df.drop(columns="Rk")

    # 3) Extract links
    soup = BeautifulSoup(html, "lxml")
    p_cells = soup.find_all("td", {"data-stat": "name_display"})
    t_cells = soup.find_all("td", {"data-stat": "team_name_abbr"})
    player_links = [BASE + c.a["href"] if c and c.a else None for c in p_cells]
    team_links = [BASE + c.a["href"] if c and c.a else None for c in t_cells]
    n = len(adv_df)
    adv_df["player_link"] = player_links[:n]
    adv_df["team_link"] = team_links[:n]
    adv_df["season"] = season

    # 4) Rename headers (handle both "Tm" and "Team")
    rename_map = {
        "Player":      "player",
        "name_display": "player",    # in case pandas read it differently
        "Age":         "age",
        "Tm":          "team",
        "Team":        "team",
        "team_name_abbr": "team",    # fallback
        "Pos":         "position",
        "G":           "games",
        "GS":          "games_started",
        "MP":          "minutes_played",
        "PER":         "per",
        "TS%":         "ts_pct",
        "3PAr":        "fg3a_per_fga_pct",
        "FTr":         "fta_per_fga_pct",
        "ORB%":        "orb_pct",
        "DRB%":        "drb_pct",
        "TRB%":        "trb_pct",
        "AST%":        "ast_pct",
        "STL%":        "stl_pct",
        "BLK%":        "blk_pct",
        "TOV%":        "tov_pct",
        "USG%":        "usg_pct",
        "OWS":         "ows",
        "DWS":         "dws",
        "WS":          "ws",
        "WS/48":       "ws_per_48",
        "OBPM":        "obpm",
        "DBPM":        "dbpm",
        "BPM":         "bpm",
        "VORP":        "vorp",
        "Awards":      "awards"
    }
    adv_df = adv_df.rename(columns=rename_map)

    # 5) Reorder to your exact spec
    order = [
        "season",
        "player", "player_link",
        "age", "team", "team_link", "position",
        "games", "games_started", "minutes_played",
        "per", "ts_pct", "fg3a_per_fga_pct", "fta_per_fga_pct",
        "orb_pct", "drb_pct", "trb_pct", "ast_pct", "stl_pct", "blk_pct",
        "tov_pct", "usg_pct", "ows", "dws", "ws", "ws_per_48",
        "obpm", "dbpm", "bpm", "vorp", "awards"
    ]
    cols = [c for c in order if c in adv_df.columns]
    return adv_df[cols]


def scrape_all():
    all_seasons = []
    for yr in range(START, END+1):
        path = f"/leagues/NBA_{yr}_advanced.html"
        print(f"Fetching {yr}…", end="")
        html = fetch(path)
        if not html:
            print(" skipped")
            continue
        df = parse_season(html, yr)
        print(f" {len(df)} rows")
        all_seasons.append(df)
    return pd.concat(all_seasons, ignore_index=True) if all_seasons else pd.DataFrame()


if __name__ == "__main__":
    result = scrape_all()
    if result.empty:
        print("No data retrieved.")
    else:
        text_cols = {"player", "player_link", "team",
                     "team_link", "position", "awards", "season"}
        for col in result.columns:
            if col in text_cols:
                continue
            result[col] = pd.to_numeric(result[col].astype(
                str).str.replace(",", ""), errors="coerce")
        result.to_csv(OUT_CSV, index=False)
        print(f"\nSaved {len(result):,} rows to {OUT_CSV}")
