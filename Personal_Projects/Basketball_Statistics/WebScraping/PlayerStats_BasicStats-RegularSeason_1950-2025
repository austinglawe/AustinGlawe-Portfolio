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
USER_AGENT = "MyTotalsScraper/1.0 (+https://example.com/info)"
START_YEAR = 1950
END_YEAR = 2025
OUTPUT_CSV = f"nba_season_totals_{START_YEAR}_{END_YEAR}.csv"

# --- Read robots.txt for crawl-delay ---
rp = robotparser.RobotFileParser()
rp.set_url(ROBOTS_URL)
rp.read()
CRAWL_DELAY = rp.crawl_delay(USER_AGENT) or 3


def fetch_url(path: str) -> str | None:
    """Fetch HTML if allowed by robots.txt and HTTP 200, else None."""
    full = BASE_URL + path
    if not rp.can_fetch(USER_AGENT, path):
        print(f"[robots.txt] Disallowed: {path}")
        return None
    time.sleep(CRAWL_DELAY + random.random())
    try:
        r = requests.get(full, headers={"User-Agent": USER_AGENT})
        if r.status_code == 200:
            return r.text
        else:
            print(f"[Warning] {r.status_code} fetching {path}")
    except Exception as e:
        print(f"[Error] fetching {path}: {e}")
    return None


def parse_totals_table(html: str, season: int) -> pd.DataFrame:
    """
    Parse /leagues/NBA_{season}_totals.html totals_stats table,
    extracting exactly the columns you requested by data-stat,
    plus player and team links.
    """
    soup = BeautifulSoup(html, "lxml")

    # 1) Find the table (may be inside HTML comments)
    table = None
    for c in soup.find_all(string=lambda t: isinstance(t, Comment)):
        if "id=\"totals_stats\"" in c:
            sub = BeautifulSoup(c, "lxml")
            table = sub.find("table", id="totals_stats")
            if table:
                break
    if table is None:
        table = soup.find("table", id="totals_stats")
    if table is None:
        print(f"[parse] No totals table for {season}")
        return pd.DataFrame()

    # 2) Extract rows
    rows = []
    for tr in table.find("tbody").find_all("tr"):
        if tr.get("class") and "thead" in tr.get("class"):
            continue
        data = {}
        for cell in tr.find_all(["th", "td"]):
            stat = cell.get("data-stat")
            if not stat:
                continue
            text = cell.get_text(strip=True)
            data[stat] = text
            # capture links
            a = cell.find("a", href=True)
            if stat == "name_display":
                data["player_link"] = BASE_URL + \
                    a["href"] if a and a["href"].startswith(
                        "/") else (a["href"] if a else None)
            elif stat == "team_name_abbr":
                data["team_link"] = BASE_URL + \
                    a["href"] if a and a["href"].startswith(
                        "/") else (a["href"] if a else None)
        data["season"] = season
        rows.append(data)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    # 3) Select & rename only requested columns
    rename_map = {
        "name_display":    "player",
        "player_link":     "player_link",
        "age":             "age",
        "team_name_abbr":  "team",
        "team_link":       "team_link",
        "pos":             "position",
        "games":           "games",
        "games_started":   "games_started",
        "mp":              "minutes_played",
        "fg":              "fgm",
        "fga":             "fga",
        "fg_pct":          "fg_pct",
        "fg3":             "fg3m",
        "fg3a":            "fg3a",
        "fg3_pct":         "fg3_pct",
        "fg2":             "fg2m",
        "fg2a":            "fg2a",
        "fg2_pct":         "fg2_pct",
        "efg_pct":         "efg_pct",
        "ft":              "ftm",
        "fta":             "fta",
        "ft_pct":          "ft_pct",
        "orb":             "oreb",
        "drb":             "dreb",
        "trb":             "treb",
        "ast":             "ast",
        "stl":             "stl",
        "blk":             "blk",
        "tov":             "tov",
        "pf":              "pf",
        "pts":             "pts",
        "tpl_dbl":         "tpl_dbl",
        "awards":          "awards",
        "season":          "season"
    }
    # filter to present keys
    cols = [k for k in rename_map if k in df.columns]
    df = df[cols].rename(columns=rename_map)

    return df


def scrape_all_seasons(start: int, end: int) -> pd.DataFrame:
    all_dfs = []
    for yr in range(start, end+1):
        path = f"/leagues/NBA_{yr}_totals.html"
        print(f"Fetching totals for {yr} …", end="")
        html = fetch_url(path)
        if not html:
            print(" skipped")
            continue
        season_df = parse_totals_table(html, yr)
        if season_df.empty:
            print(" no data")
        else:
            print(f" {len(season_df):,} rows")
            all_dfs.append(season_df)
    return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()


if __name__ == "__main__":
    totals_df = scrape_all_seasons(START_YEAR, END_YEAR)
    if totals_df.empty:
        print("No totals data retrieved.")
        exit()

    # 4) Convert numeric columns safely
    nonnum = {"player", "player_link", "team",
              "team_link", "position", "awards", "season"}
    for c in totals_df.columns:
        if c in nonnum:
            continue
        ser = totals_df[c]
        if ser.dtype == object:
            ser = ser.str.replace(",", "", regex=False)
        totals_df[c] = pd.to_numeric(ser, errors="coerce")

    # 5) Reorder columns
    final_cols = [
        "season",
        "player", "player_link",
        "age",
        "team", "team_link",
        "position",
        "games", "games_started", "minutes_played",
        "fgm", "fga", "fg_pct",
        "fg3m", "fg3a", "fg3_pct",
        "fg2m", "fg2a", "fg2_pct",
        "efg_pct",
        "ftm", "fta", "ft_pct",
        "oreb", "dreb", "treb",
        "ast", "stl", "blk", "tov", "pf", "pts",
        "tpl_dbl", "awards"
    ]
    final_cols = [c for c in final_cols if c in totals_df.columns]
    totals_df = totals_df[final_cols]

    # 6) Save to CSV
    totals_df.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSaved {len(totals_df):,} rows to {OUTPUT_CSV}")
    try:
        os.startfile(OUTPUT_CSV)
    except Exception:
        pass
