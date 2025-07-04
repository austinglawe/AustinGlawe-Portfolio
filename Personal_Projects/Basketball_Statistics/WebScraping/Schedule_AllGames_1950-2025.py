import time
import random
import requests
from urllib import robotparser
from bs4 import BeautifulSoup, Comment
import pandas as pd
import os

# --- Config ---
BASE_URL = "https://www.basketball-reference.com"
ROBOTS_URL = BASE_URL + "/robots.txt"
USER_AGENT = "MyBasketballScraper/5.0 (+https://example.com/info)"

START_YEAR = 1950
END_YEAR = 2025
MONTHS = ["august", "september", "october", "november", "december", "january", "february",
          "march", "april", "may", "june", "july"]
OUTPUT_CSV = f"nba_schedules_{START_YEAR}_{END_YEAR}.csv"

# --- robots.txt parsing ---
rp = robotparser.RobotFileParser()
rp.set_url(ROBOTS_URL)
rp.read()
crawl_delay = rp.crawl_delay(USER_AGENT) or 3


def fetch_url(path):
    full = BASE_URL + path
    if not rp.can_fetch(USER_AGENT, path):
        return None
    time.sleep(crawl_delay + random.random())
    try:
        r = requests.get(full, headers={"User-Agent": USER_AGENT})
    except Exception as e:
        print(f"[Error] fetching {full}: {e}")
        return None
    if r.status_code == 200:
        return r.text
    return None


def parse_month_schedule(html, season, month):
    """
    Parse a single month page, keyed by data-stat.
    Extracts:
      - date_game → date and date_link
      - visitor_team_name → visitor and visitor_link
      - home_team_name    → home and home_link
      - visitor_pts, home_pts, box_score_text → box_score_link
      - other stats by data-stat
    """
    soup = BeautifulSoup(html, "lxml")
    # collect tables (comment-wrapped first, then direct)
    tables = []
    for c in soup.find_all(string=lambda t: isinstance(t, Comment)):
        if "<table" in c:
            sub = BeautifulSoup(c, "lxml")
            tables += sub.find_all("table", id="games")
            tables += sub.find_all("table",
                                   attrs={"class": lambda x: x and "stats_table" in x})
    tables += soup.find_all("table", id="games")
    tables += soup.find_all("table",
                            attrs={"class": lambda x: x and "stats_table" in x})

    for tbl in tables:
        tbody = tbl.find("tbody")
        if not tbody:
            continue

        rows = []
        for tr in tbody.find_all("tr"):
            if "thead" in (tr.get("class") or []):
                continue

            cells = tr.find_all(["th", "td"])
            row = {}

            for cell in cells:
                stat = cell.get("data-stat")
                if not stat:
                    continue

                text = cell.get_text(strip=True)
                row[stat] = text

                # For any cell with a link, capture it under a new key:
                a = cell.find("a", href=True)
                if a:
                    href = a["href"]
                    full = BASE_URL + href if href.startswith("/") else href
                    if stat == "date_game":
                        row["date_link"] = full
                    elif stat == "visitor_team_name":
                        row["visitor_link"] = full
                    elif stat == "home_team_name":
                        row["home_link"] = full
                    elif stat == "box_score_text":
                        row["box_score_link"] = full

            # ensure all link keys exist even if missing
            for key in ("date_link", "visitor_link", "home_link", "box_score_link"):
                row.setdefault(key, None)

            # add season & month
            row["season"] = season
            row["month"] = month

            rows.append(row)

        if rows:
            return pd.DataFrame(rows)

    return pd.DataFrame()


def fetch_all_schedules(start_year, end_year):
    all_dfs = []
    for yr in range(start_year, end_year+1):
        for m in MONTHS:
            path = f"/leagues/NBA_{yr}_games-{m}.html"
            print(f"Fetching {path} ...", end="")
            html = fetch_url(path)
            if not html:
                print(" skipped")
                continue
            df = parse_month_schedule(html, yr, m)
            if df.empty:
                print(" no data")
            else:
                print(f" {len(df)} rows")
                all_dfs.append(df)
    return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()


if __name__ == "__main__":
    df = fetch_all_schedules(START_YEAR, END_YEAR)
    if df.empty:
        print("No data retrieved.")
        exit()

    # Rename data-stat columns to friendlier names
    rename_map = {
        "date_game":        "date",
        "game_start_time":  "start_et",
        "visitor_team_name": "visitor",
        "visitor_pts":      "visitor_pts",
        "home_team_name":   "home",
        "home_pts":         "home_pts",
        "overtimes":        "ot",
        "attendance":       "attendance",
        "game_duration":    "duration",
        "arena_name":       "arena",
        "game_remarks":     "notes"
    }
    df = df.rename(
        columns={k: v for k, v in rename_map.items() if k in df.columns})

    # Convert types
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    for col in ("visitor_pts", "home_pts", "attendance"):
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].str.replace(",", ""), errors="coerce")

    # Preview
    print("Columns:", df.columns.tolist())
    print(df.head().to_string(index=False))

    # Save CSV
    df.to_csv(OUTPUT_CSV, index=False)
    print(f"Saved to {OUTPUT_CSV}")
    try:
        os.startfile(OUTPUT_CSV)
    except:
        pass
