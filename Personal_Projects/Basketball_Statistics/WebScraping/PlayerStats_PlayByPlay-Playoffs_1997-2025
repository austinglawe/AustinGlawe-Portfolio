import time
import random
import requests
from urllib import robotparser
from bs4 import BeautifulSoup, Comment
import pandas as pd

# --- Configuration ---
BASE_URL = "https://www.basketball-reference.com"
ROBOTS_URL = BASE_URL + "/robots.txt"
USER_AGENT = "PBP_Playoff_Totals/1.0 (+https://example.com/info)"
START_YEAR = 1997
END_YEAR = 2025
OUTPUT_CSV = f"nba_playoff_pbp_totals_{START_YEAR}_{END_YEAR}.csv"

# robots.txt parser
rp = robotparser.RobotFileParser()
rp.set_url(ROBOTS_URL)
rp.read()
CRAWL_DELAY = rp.crawl_delay(USER_AGENT) or 3


def fetch_html(path: str) -> str | None:
    url = BASE_URL + path
    if not rp.can_fetch(USER_AGENT, path):
        print(f"[robots.txt] Disallowed {path}")
        return None
    time.sleep(CRAWL_DELAY + random.random())
    r = requests.get(url, headers={"User-Agent": USER_AGENT})
    if r.status_code != 200:
        print(f"[Error] HTTP {r.status_code} at {path}")
        return None
    return r.text


def parse_playoff_pbp_totals(html: str, season: int) -> pd.DataFrame:
    soup = BeautifulSoup(html, "lxml")
    table = None

    # 1) Try commented-out blocks
    for comment in soup.find_all(string=lambda t: isinstance(t, Comment)):
        if 'data-stat="pct_1"' in comment:
            tbl = BeautifulSoup(comment, "lxml").find(
                "table", class_="stats_table")
            if tbl:
                table = tbl
                break

    # 2) Fallback to any stats_table with pct_1 header
    if table is None:
        for tbl in soup.find_all("table", class_="stats_table"):
            hdrs = [th.get("data-stat", "") for th in tbl.thead.find_all("th")]
            if "pct_1" in hdrs:
                table = tbl
                break

    if table is None:
        print(f"[warn] No playoff PBP totals for {season}")
        return pd.DataFrame()

    # 3) Header order from data-stat
    headers = [th["data-stat"]
               for th in table.thead.find_all("th") if th.get("data-stat")]

    rows = []
    for tr in table.tbody.find_all("tr", class_="full_table"):
        data = {}
        # grab each cell by data-stat
        for cell in tr.find_all(["th", "td"]):
            stat = cell.get("data-stat")
            if stat not in headers:
                continue
            text = cell.get_text(strip=True)
            data[stat] = text

            # player link
            if stat == "player":
                a = cell.find("a", href=True)
                if a:
                    data["player_link"] = BASE_URL + a["href"]

            # team + link
            if stat == "team_id":
                a = cell.find("a", href=True)
                if a:
                    data["team_link"] = BASE_URL + a["href"]
                # we'll rename team_id → team

        data["season"] = season
        rows.append(data)

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    # 4) Build desired column order
    basic = ["season", "player", "player_link", "pos",
             "age", "team_id", "team_link", "g", "mp"]
    desired = basic + [
        "pct_1", "pct_2", "pct_3", "pct_4", "pct_5",
        "plus_minus_on", "plus_minus_net",
        "tov_bad_pass", "tov_lost_ball",
        "fouls_shooting", "fouls_offensive",
        "drawn_shooting", "drawn_offensive",
        "astd_pts", "and1s", "own_shots_blk"
    ]
    ordered = [c for c in desired if c in df.columns]
    df = df[ordered]

    # 5) Rename to friendly names
    df = df.rename(columns={
        "team_id":      "team",
        "player":       "player",
    })
    return df


def scrape_all_playoff_totals():
    all_seasons = []
    for year in range(START_YEAR, END_YEAR+1):
        path = f"/playoffs/NBA_{year}_play-by-play.html"
        print(f"Fetching playoffs {year} …", end="")
        html = fetch_html(path)
        if not html:
            print(" skipped")
            continue
        season_df = parse_playoff_pbp_totals(html, year)
        if season_df.empty:
            print(" no data")
        else:
            print(f" {len(season_df)} rows")
            all_seasons.append(season_df)
    return pd.concat(all_seasons, ignore_index=True) if all_seasons else pd.DataFrame()


if __name__ == "__main__":
    result = scrape_all_playoff_totals()
    if result.empty:
        print("No data retrieved.")
        exit()

    # 6) Convert numeric columns
    text_cols = {"season", "player", "player_link", "team", "team_link", "pos"}
    for col in result.columns:
        if col in text_cols:
            continue
        result[col] = pd.to_numeric(result[col].str.replace(
            "%", "").str.replace(",", ""), errors="coerce")

    result.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSaved {len(result):,} rows to {OUTPUT_CSV}")
