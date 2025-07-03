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
USER_AGENT = "MyPlayoffsScraper/2.4 (+https://example.com/info)"

START_YEAR = 1950
END_YEAR = 2025
OUTPUT_CSV = f"nba_playoff_totals_{START_YEAR}_{END_YEAR}.csv"

# Read robots.txt for crawl-delay
rp = robotparser.RobotFileParser()
rp.set_url(ROBOTS_URL)
rp.read()
CRAWL_DELAY = rp.crawl_delay(USER_AGENT) or 3


def fetch_url(path: str) -> str | None:
    full = BASE_URL + path
    if not rp.can_fetch(USER_AGENT, path):
        return None
    time.sleep(CRAWL_DELAY + random.random())
    try:
        r = requests.get(full, headers={"User-Agent": USER_AGENT})
        if r.status_code == 200:
            return r.text
    except:
        pass
    return None


def parse_playoff_totals(html: str, season: int) -> pd.DataFrame:
    soup = BeautifulSoup(html, "lxml")

    # 1) Find the totals table
    table = None
    for c in soup.find_all(string=lambda t: isinstance(t, Comment)):
        if 'id="totals_stats"' in c:
            sub = BeautifulSoup(c, "lxml")
            table = sub.find("table", id="totals_stats")
            if table:
                break
    if table is None:
        table = soup.find("table", id="totals_stats")
    if table is None:
        print(f"[parse] No playoff totals for {season}")
        return pd.DataFrame()

    # 2) Extract each row
    rows = []
    for tr in table.tbody.find_all("tr"):
        if tr.get("class") and "thead" in tr.get("class"):
            continue
        data = {}
        for cell in tr.find_all(["th", "td"]):
            stat = cell.get("data-stat")
            if not stat:
                continue
            text = cell.get_text(strip=True)
            data[stat] = text

            # player link
            if stat == "player":
                a = cell.find("a", href=True)
                data["player_link"] = BASE_URL + \
                    a["href"] if a and a["href"].startswith(
                        "/") else (a["href"] if a else None)
            # team link: now data-stat="team_id"
            if stat == "team_id":
                a = cell.find("a", href=True)
                data["team_link"] = BASE_URL + \
                    a["href"] if a and a["href"].startswith(
                        "/") else (a["href"] if a else None)

        data["season"] = season
        rows.append(data)

    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)

    # 3) Keep & rename columns exactly as before
    desired = [
        "season",
        "player", "player_link",
        "age",
        "team_id", "team_link",
        "pos",
        "games", "games_started", "mp",
        "fg", "fga", "fg_pct",
        "fg3", "fg3a", "fg3_pct",
        "fg2", "fg2a", "fg2_pct",
        "efg_pct",
        "ft", "fta", "ft_pct",
        "orb", "drb", "trb",
        "ast", "stl", "blk", "tov", "pf", "pts",
        "tpl_dbl", "awards"
    ]
    present = [c for c in desired if c in df.columns]
    df = df[present]

    rename_map = {
        "season":   "season",
        "player":   "player",
        "player_link": "player_link",
        "age":      "age",
        "team_id":  "team",
        "team_link": "team_link",
        "pos":      "position",
        "games":    "games",
        "games_started": "games_started",
        "mp":       "minutes_played",
        "fg":       "fgm",
        "fga":      "fga",
        "fg_pct":   "fg_pct",
        "fg3":      "fg3m",
        "fg3a":     "fg3a",
        "fg3_pct":  "fg3_pct",
        "fg2":      "fg2m",
        "fg2a":     "fg2a",
        "fg2_pct":  "fg2_pct",
        "efg_pct":  "efg_pct",
        "ft":       "ftm",
        "fta":      "fta",
        "ft_pct":   "ft_pct",
        "orb":      "oreb",
        "drb":      "dreb",
        "trb":      "treb",
        "ast":      "ast",
        "stl":      "stl",
        "blk":      "blk",
        "tov":      "tov",
        "pf":       "pf",
        "pts":      "pts",
        "tpl_dbl":  "tpl_dbl",
        "awards":   "awards"
    }
    df = df.rename(columns=rename_map)
    return df


def scrape_playoff_totals(start: int, end: int) -> pd.DataFrame:
    all_dfs = []
    for yr in range(start, end+1):
        path = f"/playoffs/NBA_{yr}_totals.html"
        print(f"Fetching playoffs {yr} …", end="")
        html = fetch_url(path)
        if not html:
            print(" skipped")
            continue
        season_df = parse_playoff_totals(html, yr)
        if season_df.empty:
            print(" no data")
        else:
            print(f" {len(season_df):,} rows")
            all_dfs.append(season_df)
    return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()


if __name__ == "__main__":
    df = scrape_playoff_totals(START_YEAR, END_YEAR)
    if df.empty:
        print("No playoff totals retrieved.")
        exit()

    nonnum = {"player", "player_link", "team",
              "team_link", "position", "awards", "season"}
    for c in df.columns:
        if c in nonnum:
            continue
        ser = df[c]
        if ser.dtype == object:
            ser = ser.str.replace(",", "", regex=False)
        df[c] = pd.to_numeric(ser, errors="coerce")

    df.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSaved {len(df):,} rows to {OUTPUT_CSV}")
    try:
        os.startfile(OUTPUT_CSV)
    except:
        pass
