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
USER_AGENT = "PBP_Totals_Scraper/1.2 (+https://example.com/info)"
START_YEAR = 1997
END_YEAR = 2025
OUTPUT_CSV = f"nba_pbp_totals_{START_YEAR}_{END_YEAR}.csv"

# Parse robots.txt for crawl-delay
rp = robotparser.RobotFileParser()
rp.set_url(ROBOTS_URL)
rp.read()
CRAWL_DELAY = rp.crawl_delay(USER_AGENT) or 3


def fetch_html(path: str) -> str | None:
    full = BASE_URL + path
    if not rp.can_fetch(USER_AGENT, path):
        print(f"[robots.txt] Disallowed: {path}")
        return None
    time.sleep(CRAWL_DELAY + random.random())
    resp = requests.get(full, headers={"User-Agent": USER_AGENT})
    if resp.status_code != 200:
        print(f"[Error] HTTP {resp.status_code} at {path}")
        return None
    return resp.text


def parse_pbp_totals(html: str, season: int) -> pd.DataFrame:
    soup = BeautifulSoup(html, "lxml")
    table = None

    # 1) Try inside commented blocks
    for c in soup.find_all(string=lambda t: isinstance(t, Comment)):
        if 'data-stat="pct_1"' in c:
            tbl = BeautifulSoup(c, "lxml").find("table", class_="stats_table")
            if tbl:
                table = tbl
                break

    # 2) Fallback to any stats_table that has pct_1 in its headers
    if table is None:
        for tbl in soup.find_all("table", class_="stats_table"):
            hdrs = [th.get("data-stat", "") for th in tbl.thead.find_all("th")]
            if "pct_1" in hdrs:
                table = tbl
                break

    if table is None:
        print(f"[warn] No PBP totals table for {season}")
        return pd.DataFrame()

    # 3) Gather the data-stat columns in header order
    headers = [th["data-stat"]
               for th in table.thead.find_all("th") if th.get("data-stat")]

    rows = []
    for tr in table.tbody.find_all("tr"):
        if tr.get("class") and "thead" in tr.get("class"):
            continue
        data = {}
        for cell in tr.find_all(["th", "td"]):
            stat = cell.get("data-stat")
            if stat in headers:
                data[stat] = cell.get_text(strip=True)
                if stat == "name_display":
                    a = cell.find("a", href=True)
                    if a:
                        data["player_link"] = BASE_URL + a["href"]
                if stat == "team_name_abbr":
                    a = cell.find("a", href=True)
                    if a:
                        data["team_link"] = BASE_URL + a["href"]
        data["season"] = season
        rows.append(data)

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    # 4) Build desired column order, then intersect with actual df.columns
    basic = ["season", "name_display", "player_link",
             "age", "team_name_abbr", "team_link", "pos"]
    desired = basic + [h for h in headers if h not in basic]
    # only keep those columns that made it into the DataFrame
    ordered = [c for c in desired if c in df.columns]
    df = df[ordered]

    # 5) Friendly rename
    rename_map = {
        "name_display":    "player",
        "team_name_abbr":  "team"
    }
    df = df.rename(columns=rename_map)
    return df


def scrape_all():
    all_dfs = []
    for year in range(START_YEAR, END_YEAR+1):
        path = f"/leagues/NBA_{year}_play-by-play.html"
        print(f"Fetching {year} …", end="")
        html = fetch_html(path)
        if not html:
            print(" skipped")
            continue
        season_df = parse_pbp_totals(html, year)
        if season_df.empty:
            print(" no data")
        else:
            print(f" {len(season_df)} rows")
            all_dfs.append(season_df)
    return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()


if __name__ == "__main__":
    result = scrape_all()
    if result.empty:
        print("No data retrieved.")
        exit()

    # 6) Convert numeric columns except these
    text_cols = {"season", "player", "player_link", "team", "team_link", "pos"}
    for col in result.columns:
        if col in text_cols:
            continue
        result[col] = pd.to_numeric(
            result[col].str.replace(",", ""), errors="coerce")

    result.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSaved {len(result):,} rows to {OUTPUT_CSV}")
