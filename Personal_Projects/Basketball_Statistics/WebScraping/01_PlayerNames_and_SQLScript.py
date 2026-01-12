# --------------------------------------------------------------
# IMPORTS — WHAT THEY ARE AND WHY WE NEED THEM
# --------------------------------------------------------------

import random
# Used to generate random wait times between requests.
# This prevents sending perfectly timed, bot-like requests.

import time
# Provides time.sleep() for polite delays between page loads.

from datetime import datetime
# Used to generate a timestamp for output filenames.

from zoneinfo import ZoneInfo
# Ensures timestamps always use a consistent timezone ("America/Chicago").

import pandas as pd
# Core data structure handling (DataFrame) and Excel export.

from bs4 import BeautifulSoup
# Parses HTML returned from Basketball-Reference.

from playwright.sync_api import sync_playwright
# Controls a real browser (Chromium) to load dynamic BR content.

from openpyxl.utils import get_column_letter
# Used for Excel column autofit functionality.

from openpyxl.styles import Alignment
# Allows us to left-align header cells.

from openpyxl.worksheet.views import Selection
# Used to select cell A1 by default when opening Excel.

import os
# Used for safe directory creation if needed.


# --------------------------------------------------------------
# GLOBAL SETTINGS — USER CONFIGURATIONS
# --------------------------------------------------------------

BASE_URL = "https://www.basketball-reference.com"

# Letter range to scrape (script will auto-fix reversed inputs).
START_LETTER = "b"
END_LETTER = "a"

# What outputs should be generated?
INCLUDE_HTML = True           # Save raw HTML pages A-Z to a combined file.
INCLUDE_PLAYER_DATA = True    # Parse player table into Excel.
INCLUDE_SQL_SCRIPT = True     # Generate CREATE TABLE + INSERT statements.

# Delay settings (politeness to the host website).
MIN_DELAY_SECONDS = 1.0
MAX_DELAY_SECONDS = 3.0


# --------------------------------------------------------------
# LETTER NORMALIZATION — FIX INPUT RANGE
# --------------------------------------------------------------

def clamp_letters(start_letter_input: str, end_letter_input: str) -> tuple[str, str]:
    """
    Normalize letter inputs and ensure the range is valid.
    Example: if user enters B → A, this becomes A → B.
    """

    start_letter_normalized = (start_letter_input or "a").strip().lower()[:1]
    end_letter_normalized = (end_letter_input or "z").strip().lower()[:1]

    # Default invalid characters
    if not ("a" <= start_letter_normalized <= "z"):
        start_letter_normalized = "a"
    if not ("a" <= end_letter_normalized <= "z"):
        end_letter_normalized = "z"

    # Swap if reversed
    if end_letter_normalized < start_letter_normalized:
        start_letter_normalized, end_letter_normalized =
        (
            end_letter_normalized,
            start_letter_normalized,
        )

    return start_letter_normalized, end_letter_normalized


def build_letter_range_list(start_letter: str, end_letter: str) -> list[str]:
    """Return the list of letters between start and end inclusive."""
    return [chr(x) for x in range(ord(start_letter), ord(end_letter) + 1)]


# --------------------------------------------------------------
# POLITE DELAY
# --------------------------------------------------------------

def wait_politely():
    """Sleep a random duration between MIN_DELAY_SECONDS and MAX_DELAY_SECONDS."""
    delay = random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS)
    time.sleep(delay)


# --------------------------------------------------------------
# BUILD URL
# --------------------------------------------------------------

def build_players_letter_url(letter: str) -> str:
    """Return the BR URL for a players list for a given letter."""
    return f"{BASE_URL}/players/{letter}/"


# --------------------------------------------------------------
# DATA CONVERSION HELPERS
# --------------------------------------------------------------

def safe_to_float(value: str):
    """Convert to float or return None."""
    try:
        return float(value)
    except Exception:
        return None


def safe_to_int(value: str):
    """Convert to int or return None."""
    try:
        return int(value)
    except Exception:
        return None


def inches_to_meters(height_inches: float | None):
    """Convert inches → meters."""
    return round(height_inches * 0.0254, 3) if height_inches is not None else None


def pounds_to_kilograms(weight_pounds: int | None):
    """Convert pounds → kg."""
    return round(weight_pounds * 0.45359237, 3) if weight_pounds is not None else None


def normalize_height_feet_inches(height_string: str):
    """
    Convert "6-9" → "6-09".
    Ensures consistent formatting and sorting.
    """
    if not height_string or "-" not in height_string:
        return height_string
    feet, inches = height_string.split("-", 1)
    return f"{feet}-{inches.zfill(2)}"


# --------------------------------------------------------------
# PARSE THE PLAYERS TABLE HTML
# --------------------------------------------------------------

def parse_players_table_html(table_html: str, letter: str) -> list[dict]:
    """Extract player info from the players table HTML."""

    soup = BeautifulSoup(table_html, "html.parser")
    tbody = soup.find("tbody")
    if not tbody:
        return []

    extracted_rows = []

    for tr in tbody.find_all("tr", recursive=False):

        player_cell = tr.find(["th", "td"], attrs={"data-stat": "player"})
        if not player_cell:
            continue

        unique_id = player_cell.get("data-append-csv", "") or ""

        anchor_tag = player_cell.find("a")
        if not anchor_tag or not anchor_tag.get("href"):
            continue

        player_name = anchor_tag.get_text(strip=True)
        player_url = BASE_URL + anchor_tag["href"]

        hof_flag = "x" if "*" in player_cell.get_text(strip=True) else ""
        active_flag = "Active" if player_cell.find("strong") else "Inactive"

        def get_text(stat):
            td = tr.find("td", attrs={"data-stat": stat})
            return td.get_text(strip=True) if td else ""

        def get_attr(stat, attr):
            td = tr.find("td", attrs={"data-stat": stat})
            return td.get(attr) if td else None

        first_year = get_text("year_min")
        last_year = get_text("year_max")
        position = get_text("pos")

        height_ft_raw = get_text("height")
        height_ft_norm = normalize_height_feet_inches(height_ft_raw)

        height_in_raw = get_attr("height", "csk") or ""
        height_in = safe_to_float(height_in_raw)
        height_m = inches_to_meters(height_in)

        weight_lb = safe_to_int(get_text("weight"))
        weight_kg = pounds_to_kilograms(weight_lb)

        birth_csk = (get_attr("birth_date", "csk") or "").strip()
        birthday_short = (
            f"{birth_csk[0:4]}.{birth_csk[4:6]}.{birth_csk[6:8]}"
            if len(birth_csk) == 8 else ""
        )
        birthday_long = get_text("birth_date")

        colleges_cell = tr.find("td", attrs={"data-stat": "colleges"})
        colleges = ""
        if colleges_cell:
            colleges = " | ".join(a.get_text(strip=True)
                                  for a in colleges_cell.find_all("a"))

        extracted_rows.append({
            "letter": letter.upper(),
            "unique_id": unique_id,
            "player_name": player_name,
            "player_url": player_url,
            "HOF": hof_flag,
            "status": active_flag,
            "year_start": first_year,
            "year_end": last_year,
            "pos": position,
            "height_in": height_in,
            "height_ft_in": height_ft_norm,
            "height_m": height_m,
            "weight_lb": weight_lb,
            "weight_kg": weight_kg,
            "birthday": birthday_short,
            "birthday_long": birthday_long,
            "colleges": colleges,
        })

    return extracted_rows


# --------------------------------------------------------------
# HTML FILE HEADER
# --------------------------------------------------------------

def write_html_file_header(file_handle, title, start_letter, end_letter, timestamp):
    """Write descriptive header text into HTML output file."""
    file_handle.write(f"{title}\n")
    file_handle.write(f"Letters: {start_letter.upper()}-{end_letter.upper()}\n")
    file_handle.write(f"Generated: {timestamp} America/Chicago\n\n")


# --------------------------------------------------------------
# SILENT ROUTE HANDLER — NO CancelledError SPAM
# --------------------------------------------------------------

def safe_route_handler(route):
    """
    Abort images/css/fonts/media but suppress cancellation warnings.
    Ensures a clean console output.
    """
    try:
        if route.request.resource_type in ["image", "stylesheet", "font", "media"]:
            return route.abort()
        return route.continue_()
    except Exception:
        pass  # Suppress harmless async warnings


# --------------------------------------------------------------
# ESCAPING FOR SQL (APOSTROPHES)
# --------------------------------------------------------------

def escape_sql_string(value: str):
    """Escape apostrophes for SQLite by doubling them."""
    return value.replace("'", "''") if isinstance(value, str) else value


# --------------------------------------------------------------
# SQL SCRIPT GENERATION
# --------------------------------------------------------------

def generate_sql_script(
    file_path: str,
    start_letter: str,
    end_letter: str,
    timestamp_string: str,
    rows: list[dict]
):
    """
    Write out a full SQLite SQL script including:
      - Comments
      - CREATE TABLE
      - INSERT statements (one per row)
    """

    with open(file_path, "w", encoding="utf-8") as sql_file:

        # ------------------------------------------
        # HEADER COMMENTS
        # ------------------------------------------
        sql_file.write("-- ================================================\n")
        sql_file.write("-- Basketball-Reference Players A-Z SQL Export\n")
        sql_file.write(f"-- Letters: {start_letter.upper()}-{end_letter.upper()}\n")
        sql_file.write(f"-- Generated: {timestamp_string} America/Chicago\n")
        sql_file.write("-- ================================================\n\n")

        sql_file.write("-- Column Descriptions:\n")
        sql_file.write("-- unique_id: Basketball-Reference unique player id\n")
        sql_file.write("-- player_name: Player's full displayed name\n")
        sql_file.write("-- letter: A-Z grouping based on BR players list\n")
        sql_file.write("-- player_url: URL to player's profile page\n")
        sql_file.write("-- HOF: 'x' if Hall of Fame\n")
        sql_file.write("-- status: 'Active' or 'Inactive'\n")
        sql_file.write("-- year_start: First season\n")
        sql_file.write("-- year_end: Final season\n")
        sql_file.write("-- pos: Position\n")
        sql_file.write("-- height_in: Height in inches\n")
        sql_file.write("-- height_ft_in: Height in feet-inches format\n")
        sql_file.write("-- height_m: Height in meters\n")
        sql_file.write("-- weight_lb: Weight in pounds\n")
        sql_file.write("-- weight_kg: Weight in kilograms\n")
        sql_file.write("-- birthday: YYYY.MM.DD\n")
        sql_file.write("-- birthday_long: Long birth date\n")
        sql_file.write("-- colleges: Colleges attended\n\n")

        # ------------------------------------------
        # CREATE TABLE
        # ------------------------------------------
        sql_file.write("CREATE TABLE nba_players (\n")
        sql_file.write("    unique_id TEXT PRIMARY KEY,\n")
        sql_file.write("    player_name TEXT,\n")
        sql_file.write("    letter TEXT,\n")
        sql_file.write("    player_url TEXT,\n")
        sql_file.write("    HOF TEXT,\n")
        sql_file.write("    status TEXT,\n")
        sql_file.write("    year_start INTEGER,\n")
        sql_file.write("    year_end INTEGER,\n")
        sql_file.write("    pos TEXT,\n")
        sql_file.write("    height_in REAL,\n")
        sql_file.write("    height_ft_in TEXT,\n")
        sql_file.write("    height_m REAL,\n")
        sql_file.write("    weight_lb INTEGER,\n")
        sql_file.write("    weight_kg REAL,\n")
        sql_file.write("    birthday TEXT,\n")
        sql_file.write("    birthday_long TEXT,\n")
        sql_file.write("    colleges TEXT\n")
        sql_file.write(");\n\n")

        # ------------------------------------------
        # INSERT STATEMENTS
        # ------------------------------------------
        sql_file.write("-- ================================================\n")
        sql_file.write("-- INSERT PLAYER ROWS\n")
        sql_file.write("-- ================================================\n\n")

        for row in rows:

            # Escape apostrophes
            safe_row = {k: escape_sql_string(v) for k, v in row.items()}

            sql_file.write(
                "INSERT INTO nba_players (unique_id, player_name, letter, "
                "player_url, HOF, status, year_start, year_end, pos, height_in, "
                "height_ft_in, height_m, weight_lb, weight_kg, birthday, "
                "birthday_long, colleges)\n"
            )

            sql_file.write("VALUES (\n")
            sql_file.write(f"    '{safe_row['unique_id']}',\n")
            sql_file.write(f"    '{safe_row['player_name']}',\n")
            sql_file.write(f"    '{safe_row['letter']}',\n")
            sql_file.write(f"    '{safe_row['player_url']}',\n")
            sql_file.write(f"    '{safe_row['HOF']}',\n")
            sql_file.write(f"    '{safe_row['status']}',\n")
            sql_file.write(f"    {safe_row['year_start'] if safe_row['year_start'] else 'NULL'},\n")
            sql_file.write(f"    {safe_row['year_end'] if safe_row['year_end'] else 'NULL'},\n")
            sql_file.write(f"    '{safe_row['pos']}',\n")
            sql_file.write(f"    {safe_row['height_in'] if safe_row['height_in'] else 'NULL'},\n")
            sql_file.write(f"    '{safe_row['height_ft_in']}',\n")
            sql_file.write(f"    {safe_row['height_m'] if safe_row['height_m'] else 'NULL'},\n")
            sql_file.write(f"    {safe_row['weight_lb'] if safe_row['weight_lb'] else 'NULL'},\n")
            sql_file.write(f"    {safe_row['weight_kg'] if safe_row['weight_kg'] else 'NULL'},\n")
            sql_file.write(f"    '{safe_row['birthday']}',\n")
            sql_file.write(f"    '{safe_row['birthday_long']}',\n")
            sql_file.write(f"    '{safe_row['colleges']}'\n")
            sql_file.write(");\n\n")


# --------------------------------------------------------------
# MAIN PROGRAM
# --------------------------------------------------------------

def main():

    # Normalize letters
    start_letter_normalized, end_letter_normalized =
    clamp_letters(START_LETTER, END_LETTER)

    # Timestamp for filenames
    timestamp_string = datetime.now(ZoneInfo("America/Chicago")).strftime(
        "%Y.%m.%d-%H.%M.%S"
    )

    # File name base
    base_file_name = (
        f"BR Players A-Z ("
        f"{start_letter_normalized.upper()}-{end_letter_normalized.upper()}) "
        f"as of {timestamp_string}"
    )

    html_file_name = f"{base_file_name} HTML.txt"
    excel_file_name = f"{base_file_name}.xlsx"
    sql_file_name = f"{base_file_name} SQL SCRIPT.txt"

    letters = build_letter_range_list(start_letter_normalized, end_letter_normalized)
    all_rows = []

    # Start Playwright
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()

        # Silent blocking of heavy resources
        page.route("**/*", safe_route_handler)

        html_handle = None
        if INCLUDE_HTML:
            html_handle = open(html_file_name, "w", encoding="utf-8")
            write_html_file_header(
                html_handle,
                "Basketball-Reference players A-Z pages",
                start_letter_normalized,
                end_letter_normalized,
                timestamp_string,
            )

        try:
            for idx, letter in enumerate(letters, start=1):

                url = build_players_letter_url(letter)
                print(f"[{idx}/{len(letters)}] Downloading '{letter.upper()}': {url}")

                # Load page
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(500)

                # Save HTML
                if INCLUDE_HTML and html_handle:
                    html = page.content()
                    section_header = (
                        "\n" + "=" * 120 + "\n" +
                        f"BEGIN LETTER {letter.upper()} | URL {url}\n" +
                        "=" * 120 + "\n"
                    )
                    section_footer = (
                        "\n" + "-" * 120 + "\n" +
                        f"END LETTER {letter.upper()}\n" +
                        "-" * 120 + "\n"
                    )
                    html_handle.write(section_header)
                    html_handle.write(html)
                    html_handle.write(section_footer)

                # Parse data
                if INCLUDE_PLAYER_DATA:
                    page.wait_for_selector("table#players", state="attached", timeout=60000)
                    table_html = page.locator("table#players").evaluate(
                        "el => el.outerHTML")
                    rows = parse_players_table_html(table_html, letter)
                    print(f"{letter.upper()}: extracted {len(rows)} players")
                    all_rows.extend(rows)

                wait_politely()

        finally:
            if html_handle:
                html_handle.close()
            browser.close()

    # ----------------------------------------------------------
    # EXCEL OUTPUT
    # ----------------------------------------------------------
    if INCLUDE_PLAYER_DATA and all_rows:

        df = pd.DataFrame(all_rows)
        df = df.drop_duplicates(subset=["unique_id"])
        df = df.sort_values(["letter", "player_name"]).reset_index(drop=True)

        # Desired column order
        desired_cols = [
            "letter", "unique_id", "player_name", "player_url",
            "HOF", "status",
            "year_start", "year_end", "pos",
            "height_in", "height_ft_in", "height_m",
            "weight_lb", "weight_kg",
            "birthday", "birthday_long", "colleges"
        ]
        df = df[[c for c in desired_cols if c in df.columns]]

        with pd.ExcelWriter(excel_file_name, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="NBA Players")
            ws = writer.sheets["NBA Players"]

            # Filter on top row
            ws.auto_filter.ref = ws.dimensions

            # Left align headers
            for cell in ws[1]:
                cell.alignment = Alignment(horizontal="left")

            # Autofit columns
            for col_idx, col_cells in enumerate(ws.columns, start=1):
                max_len = max(
                    (len(str(cell.value)) if cell.value is not None else 0)
                    for cell in col_cells
                )
                ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2

            # Select A1 by default
            ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

        print(f"\nSaved Excel data file: {excel_file_name}")

    # ----------------------------------------------------------
    # SQL SCRIPT OUTPUT
    # ----------------------------------------------------------
    if INCLUDE_SQL_SCRIPT and all_rows:
        generate_sql_script(
            sql_file_name,
            start_letter_normalized,
            end_letter_normalized,
            timestamp_string,
            all_rows
        )
        print(f"Saved SQL script file: {sql_file_name}")

    # ----------------------------------------------------------
    # HTML MESSAGE
    # ----------------------------------------------------------
    if INCLUDE_HTML:
        print(f"Saved combined HTML file: {html_file_name}")


# --------------------------------------------------------------
# ENTRY POINT
# --------------------------------------------------------------

if __name__ == "__main__":
    main()
