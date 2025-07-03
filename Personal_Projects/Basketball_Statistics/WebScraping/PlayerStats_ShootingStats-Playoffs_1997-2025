import time
import random
import requests
from urllib import robotparser
from bs4 import BeautifulSoup, Comment
import pandas as pd

# --- Configuration ---
BASE_URL = "https://www.basketball-reference.com"
ROBOTS_URL = BASE_URL + "/robots.txt"
USER_AGENT = "PlayoffShootingScraper/1.0 (+https://example.com/info)"
START_YEAR = 1997
END_YEAR = 2025
OUTPUT_CSV = f"nba_playoff_shooting_{START_YEAR}_{END_YEAR}.csv"

# Parse robots.txt for crawl-delay
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


def parse_playoff_shooting(html: str, season: int) -> pd.DataFrame:
    soup = BeautifulSoup(html, "lxml")
    table = None

    # 1) Try inside HTML comments
    for c in soup.find_all(string=lambda t: isinstance(t, Comment)):
        if 'data-stat="fg_pct"' in c:
            tbl = BeautifulSoup(c, "lxml").find("table", class_="stats_table")
            if tbl:
                table = tbl
                break

    # 2) Fallback: any stats_table with fg_pct
    if table is None:
        for tbl in soup.find_all("table", class_="stats_table"):
            hdrs = [th.get("data-stat", "") for th in tbl.thead.find_all("th")]
            if "fg_pct" in hdrs:
                table = tbl
                break

    if table is None:
        print(f"[warn] No playoff shooting table for {season}")
        return pd.DataFrame()

    # 3) Build header list in order
    headers = [th["data-stat"]
               for th in table.thead.find_all("th") if th.get("data-stat")]

    rows = []
    for tr in table.tbody.find_all("tr", class_="full_table"):
        data = {}
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

            # team link
            if stat == "team_id":
                a = cell.find("a", href=True)
                if a:
                    data["team_link"] = BASE_URL + a["href"]

        data["season"] = season
        rows.append(data)

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    # 4) Desired column order
    basic = [
        "season",
        "player", "player_link",
        "pos", "age",
        "team_id", "team_link",
        "g", "mp"
    ]
    shooting = [
        "fg_pct", "avg_dist",
        "pct_fga_fg2a", "pct_fga_00_03", "pct_fga_03_10", "pct_fga_10_16", "pct_fga_16_xx", "pct_fga_fg3a",
        "fg_pct_fg2a", "fg_pct_00_03", "fg_pct_03_10", "fg_pct_10_16", "fg_pct_16_xx", "fg_pct_fg3a",
        "pct_ast_fg2", "pct_ast_fg3",
        "pct_fga_dunk", "fg_dunk",
        "pct_fg3a_corner3", "fg_pct_corner3",
        "fg3a_heave", "fg3_heave",
        "awards"
    ]
    desired = basic + shooting
    ordered = [c for c in desired if c in df.columns]
    df = df[ordered]

    # 5) Rename to friendly names
    df = df.rename(columns={
        "player": "player",
        "team_id": "team"
    })
    return df


def scrape_all_playoff_shooting():
    all_seasons = []
    for year in range(START_YEAR, END_YEAR+1):
        path = f"/playoffs/NBA_{year}_shooting.html"
        print(f"Fetching playoffs shooting {year} …", end="")
        html = fetch_html(path)
        if not html:
            print(" skipped")
            continue
        season_df = parse_playoff_shooting(html, year)
        if season_df.empty:
            print(" no data")
        else:
            print(f" {len(season_df)} rows")
            all_seasons.append(season_df)
    return pd.concat(all_seasons, ignore_index=True) if all_seasons else pd.DataFrame()


if __name__ == "__main__":
    result = scrape_all_playoff_shooting()
    if result.empty:
        print("No playoff shooting data retrieved.")
        exit()

    # 6) Convert numeric columns (strip % and commas)
    text_cols = {"season", "player", "player_link",
                 "team", "team_link", "pos", "awards"}
    for col in result.columns:
        if col in text_cols:
            continue
        result[col] = pd.to_numeric(result[col].str.replace(
            "%", "").str.replace(",", ""), errors="coerce")

    result.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSaved {len(result):,} rows to {OUTPUT_CSV}")
