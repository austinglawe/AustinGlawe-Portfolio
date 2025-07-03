import time
import random
import requests
from urllib import robotparser
from bs4 import BeautifulSoup, Comment
import pandas as pd
import os

# --- Configuration ---
BASE_URL = "https://www.basketball-reference.com"
ROBOTS_URL = BASE_URL + "/robots.txt"
USER_AGENT = "MyPlayoffAdvancedScraper/1.0 (+https://example.com/info)"
START_YEAR = 1950
END_YEAR = 2025
OUTPUT_CSV = f"nba_playoff_advanced_{START_YEAR}_{END_YEAR}.csv"

# Read robots.txt crawl-delay
rp = robotparser.RobotFileParser()
rp.set_url(ROBOTS_URL)
rp.read()
CRAWL_DELAY = rp.crawl_delay(USER_AGENT) or 3


def fetch_url(path: str) -> str | None:
    """Fetch page if allowed and return HTML or None."""
    url = BASE_URL + path
    if not rp.can_fetch(USER_AGENT, path):
        return None
    time.sleep(CRAWL_DELAY + random.random())
    try:
        r = requests.get(url, headers={"User-Agent": USER_AGENT})
        if r.status_code == 200:
            return r.text
    except:
        pass
    return None


def parse_playoff_advanced(html: str, season: int) -> pd.DataFrame:
    """Parse the playoff advanced_stats table for one season."""
    soup = BeautifulSoup(html, "lxml")

    # 1) Locate the advanced_stats table (may be wrapped in comments)
    table = None
    for c in soup.find_all(string=lambda t: isinstance(t, Comment)):
        if 'id="advanced_stats"' in c:
            sub = BeautifulSoup(c, "lxml")
            table = sub.find("table", id="advanced_stats")
            if table:
                break
    if table is None:
        table = soup.find("table", id="advanced_stats")
    if table is None:
        print(f"[parse] No playoff advanced table for {season}")
        return pd.DataFrame()

    # 2) Extract rows
    rows = []
    for tr in table.tbody.find_all("tr"):
        if tr.get("class") and "thead" in tr.get("class"):
            continue
        data = {}
        for cell in tr.find_all(["th", "td"]):
            stat = cell.get("data-stat")
            if not stat:
                continue
            data[stat] = cell.get_text(strip=True)
            # player link
            if stat == "player":
                a = cell.find("a", href=True)
                if a:
                    data["player_link"] = BASE_URL + a["href"]
            # team link
            if stat == "team_id":
                a = cell.find("a", href=True)
                if a:
                    data["team_link"] = BASE_URL + a["href"]
        data["season"] = season
        rows.append(data)

    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)

    # 3) Keep & rename exactly these advanced columns
    desired = [
        "season",
        "player", "player_link",
        "age",
        "team_id", "team_link",
        "pos",
        "g", "mp",
        "per", "ts_pct", "fg3a_per_fga_pct", "fta_per_fga_pct",
        "orb_pct", "drb_pct", "trb_pct",
        "ast_pct", "stl_pct", "blk_pct",
        "tov_pct", "usg_pct",
        "ows", "dws", "ws", "ws_per_48",
        "obpm", "dbpm", "bpm", "vorp"
    ]
    present = [c for c in desired if c in df.columns]
    df = df[present]

    rename_map = {
        "season":            "season",
        "player":            "player",
        "player_link":       "player_link",
        "age":               "age",
        "team_id":           "team",
        "team_link":         "team_link",
        "pos":               "position",
        "g":                 "games",
        "mp":                "minutes_played",
        "per":               "per",
        "ts_pct":            "ts_pct",
        "fg3a_per_fga_pct":  "threepar",
        "fta_per_fga_pct":   "ftr",
        "orb_pct":           "oreb_pct",
        "drb_pct":           "dreb_pct",
        "trb_pct":           "treb_pct",
        "ast_pct":           "ast_pct",
        "stl_pct":           "stl_pct",
        "blk_pct":           "blk_pct",
        "tov_pct":           "tov_pct",
        "usg_pct":           "usg_pct",
        "ows":               "ows",
        "dws":               "dws",
        "ws":                "ws",
        "ws_per_48":         "ws_per_48",
        "obpm":              "obpm",
        "dbpm":              "dbpm",
        "bpm":               "bpm",
        "vorp":              "vorp"
    }
    df = df.rename(columns=rename_map)
    return df


def scrape_playoff_advanced(start: int, end: int) -> pd.DataFrame:
    all_dfs = []
    for yr in range(start, end+1):
        path = f"/playoffs/NBA_{yr}_advanced.html"
        print(f"Fetching playoff advanced {yr} …", end="")
        html = fetch_url(path)
        if not html:
            print(" skipped")
            continue
        season_df = parse_playoff_advanced(html, yr)
        if season_df.empty:
            print(" no data")
        else:
            print(f" {len(season_df):,} rows")
            all_dfs.append(season_df)
    return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()


if __name__ == "__main__":
    adv_df = scrape_playoff_advanced(START_YEAR, END_YEAR)
    if adv_df.empty:
        print("No playoff advanced data retrieved.")
        exit()

    # Convert numeric columns
    nonnum = {"player", "player_link", "team",
              "team_link", "position", "season"}
    for c in adv_df.columns:
        if c in nonnum:
            continue
        ser = adv_df[c]
        if ser.dtype == object:
            ser = ser.str.replace(",", "", regex=False)
        adv_df[c] = pd.to_numeric(ser, errors="coerce")

    # Save to CSV
    adv_df.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSaved {len(adv_df):,} rows to {OUTPUT_CSV}")
    try:
        os.startfile(OUTPUT_CSV)
    except:
        pass
