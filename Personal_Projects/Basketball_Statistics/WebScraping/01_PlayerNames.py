# --------------------------------------------------------------
# IMPORTS — WHAT THEY ARE AND WHY WE NEED THEM
# --------------------------------------------------------------

import random
# Used to generate a random wait time between page requests so that
# our scraping does NOT appear like a bot hitting the server at
# exact intervals. This is more polite and reduces the chance of blocking.

import time
# Provides time.sleep(), which we use to delay between page loads.

from datetime import datetime
# Used to generate timestamps for file names, such as:
# "2026.01.11-19.45.22". This makes each output file unique.

from zoneinfo import ZoneInfo
# Allows us to explicitly specify a time zone, so that timestamps
# always follow the same zone ("America/Chicago") regardless of where
# the script is run.

import pandas as pd
# Provides data structures (like DataFrame) to store all scraped
# player rows and export them to Excel. Also handles sorting,
# deduplication, and column ordering.

from bs4 import BeautifulSoup
# Parses HTML into a searchable tree. Basketball-Reference player
# pages contain a hidden table inside comments; using Playwright +
# BeautifulSoup lets us reliably extract it.

from playwright.sync_api import sync_playwright
# Controls a real Chromium browser so pages load exactly like they
# would for a human. This handles dynamically inserted HTML and
# ensures the players table is present in the DOM for parsing.

from openpyxl.utils import get_column_letter
# Converts a numeric column index (1, 2, 3...) into the Excel letters
# ("A", "B", "C", ...). Needed to set autofit column widths.

from openpyxl.styles import Alignment
# Used to left-align header cells in the Excel file.

from openpyxl.worksheet.views import Selection
# Used to make Excel select cell A1 when the workbook opens.


# --------------------------------------------------------------
# TOP-LEVEL SETTINGS — USER CONTROLS
# --------------------------------------------------------------

BASE_URL = "https://www.basketball-reference.com"

# You may change these to any letters 'a'–'z'.
START_LETTER = "a"
END_LETTER = "z"

# Instead of MODE strings, we now use explicit booleans:
INCLUDE_HTML = True           # Save raw HTML pages to one text file
INCLUDE_PLAYER_DATA = True    # Parse player table and export to Excel

# Polite delay ranges between page requests
MIN_DELAY_SECONDS = 1.0
MAX_DELAY_SECONDS = 3.0


# --------------------------------------------------------------
# LETTER NORMALIZATION — HANDLE INVALID INPUT ORDERING
# --------------------------------------------------------------

def clamp_letters(start_letter_input: str, end_letter_input: str) -> tuple[str, str]:
    """
    Normalize the input letters and ensure the range is valid:
      - Forces lowercase
      - Strips whitespace
      - Defaults out-of-range values to a/z
      - Swaps letters if END < START (example: B-A becomes A-B)
    """

    start_letter_normalized = (start_letter_input or "a").strip().lower()[:1]
    end_letter_normalized = (end_letter_input or "z").strip().lower()[:1]

    # Force valid range
    if not ("a" <= start_letter_normalized <= "z"):
        start_letter_normalized = "a"
    if not ("a" <= end_letter_normalized <= "z"):
        end_letter_normalized = "z"

    # Swap if user entered them reversed
    if end_letter_normalized < start_letter_normalized:
        start_letter_normalized, end_letter_normalized = (
            end_letter_normalized,
            start_letter_normalized,
        )

    return start_letter_normalized, end_letter_normalized


def build_letter_range_list(start_letter: str, end_letter: str) -> list[str]:
    """
    Build a list like ['a', 'b', 'c'] from a letter range.
    Uses ASCII codes to generate the sequence.
    """
    return [chr(code) for code in range(ord(start_letter), ord(end_letter) + 1)]


# --------------------------------------------------------------
# DELAY HELPER
# --------------------------------------------------------------

def wait_politely():
    """
    Sleep for a random duration between MIN_DELAY_SECONDS and
    MAX_DELAY_SECONDS. This prevents overloading the server and
    reduces the chance of being rate-limited.
    """
    delay = random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS)
    time.sleep(delay)


# --------------------------------------------------------------
# URL BUILDER
# --------------------------------------------------------------

def build_players_letter_url(letter: str) -> str:
    """Return the Basketball-Reference players list URL for a letter."""
    return f"{BASE_URL}/players/{letter}/"


# --------------------------------------------------------------
# CONVERSION HELPERS — SAFE PARSING
# --------------------------------------------------------------

def safe_to_float(value: str):
    """Convert string to float or return None if invalid."""
    try:
        return float(value)
    except Exception:
        return None


def safe_to_int(value: str):
    """Convert string to int or return None if invalid."""
    try:
        return int(value)
    except Exception:
        return None


def inches_to_meters(height_inches: float | None):
    """Convert inches into meters, rounded to 3 decimals."""
    return round(height_inches * 0.0254, 3) if height_inches is not None else None


def pounds_to_kilograms(weight_pounds: int | None):
    """Convert pounds into kilograms, rounded to 3 decimals."""
    return round(weight_pounds * 0.45359237, 3) if weight_pounds is not None else None


def normalize_height_feet_inches(height_string: str):
    """
    Convert a string like "6-9" to "6-09" to keep a standard width.
    The second part gets padded with zeros if needed.
    """
    if not height_string or "-" not in height_string:
        return height_string
    feet, inches = height_string.split("-", 1)
    return f"{feet}-{inches.zfill(2)}"


# --------------------------------------------------------------
# PARSE PLAYERS TABLE HTML — EXTRACT REAL PLAYER DATA
# --------------------------------------------------------------

def parse_players_table_html(table_html: str, letter: str) -> list[dict]:
    """
    Parse the inner players table HTML (extracted with Playwright) and
    return a list of dictionaries, each representing a player's data
    for the given letter.
    """

    soup = BeautifulSoup(table_html, "html.parser")
    tbody = soup.find("tbody")
    if not tbody:
        return []

    extracted_rows = []

    # Each table row (<tr>) corresponds to a player
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

        # HOF status: "*" appears in the cell text
        hof_flag = "x" if "*" in player_cell.get_text(strip=True) else ""

        # Active = bold name
        active_flag = "Active" if player_cell.find("strong") else "Inactive"

        # Helpers to simplify extraction
        def get_text(stat):
            td = tr.find("td", attrs={"data-stat": stat})
            return td.get_text(strip=True) if td else ""

        def get_attr(stat, attr):
            td = tr.find("td", attrs={"data-stat": stat})
            return td.get(attr) if td else None

        first_year = get_text("year_min")
        last_year = get_text("year_max")
        position = get_text("pos")

        # Height
        height_ft_raw = get_text("height")
        height_ft_norm = normalize_height_feet_inches(height_ft_raw)

        height_in_raw = get_attr("height", "csk") or ""
        height_in = safe_to_float(height_in_raw)
        height_m = inches_to_meters(height_in)

        # Weight
        weight_lb = safe_to_int(get_text("weight"))
        weight_kg = pounds_to_kilograms(weight_lb)

        # Birthdates
        birth_csk = (get_attr("birth_date", "csk") or "").strip()
        birthday_short = (
            f"{birth_csk[0:4]}.{birth_csk[4:6]}.{birth_csk[6:8]}"
            if len(birth_csk) == 8 else ""
        )
        birthday_long = get_text("birth_date")

        # Colleges
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
# HTML FILE HEADER (for the combined raw pages file)
# --------------------------------------------------------------

def write_html_file_header(file_handle, title, start_letter, end_letter, timestamp):
    """
    Writes a descriptive header at the top of the combined HTML output file.
    Includes:
      - Title of the scrape
      - Letters range
      - Timestamp
    """
    file_handle.write(f"{title}\n")
    file_handle.write(
        f"Letters: {start_letter.upper()}-{end_letter.upper()}\n")
    file_handle.write(f"Generated: {timestamp} America/Chicago\n\n")


# --------------------------------------------------------------
# SAFE ROUTE HANDLER — PREVENTS CancelledError SPAM
# --------------------------------------------------------------

def safe_route_handler(route):
    """
    Basketball-Reference pages load many images, CSS, fonts,
    and other resources we do not need. We block them by calling
    route.abort().

    In Playwright, aborting a route can cause internal async tasks
    to be canceled. Without this wrapper, those cancellations produce
    harmless but annoying console warnings.

    This wrapper catches and suppresses those warnings so the console
    stays clean.
    """
    try:
        if route.request.resource_type in ["image", "stylesheet", "font", "media"]:
            return route.abort()
        return route.continue_()
    except Exception:
        # Ignore harmless cancellation warnings
        pass


# --------------------------------------------------------------
# MAIN PROGRAM
# --------------------------------------------------------------

def main():

    # Normalize letters and correct reverse ordering if necessary.
    start_letter_normalized, end_letter_normalized = clamp_letters(
        START_LETTER, END_LETTER
    )

    # Create a timestamp for file names.
    timestamp_string = datetime.now(ZoneInfo("America/Chicago")).strftime(
        "%Y.%m.%d-%H.%M.%S"
    )

    # Base file name shared by both HTML and Excel outputs.
    base_file_name = (
        f"BR Players A-Z ("
        f"{start_letter_normalized.upper()}-{end_letter_normalized.upper()}) "
        f"as of {timestamp_string}"
    )

    html_file_name = f"{base_file_name} HTML.txt"
    excel_file_name = f"{base_file_name}.xlsx"

    # Build list of letters between start and end.
    letter_list = build_letter_range_list(
        start_letter_normalized, end_letter_normalized)

    # Will hold all extracted rows from all letters.
    all_rows = []

    # Start Playwright browser session.
    with sync_playwright() as pw:

        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()

        # Install silent route handler to block images/resources.
        page.route("**/*", safe_route_handler)

        # Open HTML output file if this mode is enabled.
        html_file_handle = None
        if INCLUDE_HTML:
            html_file_handle = open(html_file_name, "w", encoding="utf-8")
            write_html_file_header(
                html_file_handle,
                "Basketball-Reference players A-Z pages",
                start_letter_normalized,
                end_letter_normalized,
                timestamp_string,
            )

        try:
            # Loop through each letter page.
            for index, letter in enumerate(letter_list, start=1):

                url = build_players_letter_url(letter)
                print(
                    f"[{index}/{len(letter_list)}] Downloading '{letter.upper()}': {url}")

                # Load page fully
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(500)  # slight extra wait

                # ------------------------------------------------
                # SAVE RAW HTML (if enabled)
                # ------------------------------------------------
                if INCLUDE_HTML and html_file_handle is not None:
                    full_page_html = page.content()

                    section_header = (
                        "\n"
                        + "=" * 120 + "\n"
                        + f"BEGIN LETTER {letter.upper()} | URL {url}\n"
                        + "=" * 120 + "\n"
                    )
                    section_footer = (
                        "\n"
                        + "-" * 120 + "\n"
                        + f"END LETTER {letter.upper()}\n"
                        + "-" * 120 + "\n"
                    )

                    html_file_handle.write(section_header)
                    html_file_handle.write(full_page_html)
                    html_file_handle.write(section_footer)

                # ------------------------------------------------
                # PARSE PLAYER DATA (if enabled)
                # ------------------------------------------------
                if INCLUDE_PLAYER_DATA:

                    # Ensure table appears in DOM
                    page.wait_for_selector(
                        "table#players", state="attached", timeout=60000)

                    # Extract only the table HTML for parsing
                    table_html = page.locator("table#players").evaluate(
                        "element => element.outerHTML"
                    )

                    rows = parse_players_table_html(table_html, letter)
                    print(f"{letter.upper()}: extracted {len(rows)} players")

                    all_rows.extend(rows)

                wait_politely()

        finally:
            # Close HTML file if open
            if html_file_handle:
                html_file_handle.close()
            # Close browser regardless of errors
            browser.close()

    # --------------------------------------------------------------
    # EXCEL OUTPUT (IF ENABLED)
    # --------------------------------------------------------------

    if INCLUDE_PLAYER_DATA and all_rows:

        df = pd.DataFrame(all_rows)

        # Remove duplicate players by unique_id
        df = df.drop_duplicates(subset=["unique_id"]).reset_index(drop=True)

        # Sort by letter then name
        df = df.sort_values(["letter", "player_name"]).reset_index(drop=True)

        # Desired Excel column order
        desired_order = [
            "letter", "unique_id", "player_name", "player_url",
            "HOF", "status", "year_start", "year_end", "pos",
            "height_in", "height_ft_in", "height_m",
            "weight_lb", "weight_kg",
            "birthday", "birthday_long", "colleges",
        ]
        df = df[[col for col in desired_order if col in df.columns]]

        # Write Excel file using openpyxl so we can format it.
        with pd.ExcelWriter(excel_file_name, engine="openpyxl") as writer:

            df.to_excel(writer, index=False, sheet_name="NBA Players")

            worksheet = writer.sheets["NBA Players"]

            # --------------------------------------------------
            # ADD FILTER TO HEADER ROW
            # --------------------------------------------------
            worksheet.auto_filter.ref = worksheet.dimensions

            # --------------------------------------------------
            # LEFT-ALIGN HEADER CELLS
            # --------------------------------------------------
            for cell in worksheet[1]:
                cell.alignment = Alignment(horizontal="left")

            # --------------------------------------------------
            # AUTOFIT EACH COLUMN
            # --------------------------------------------------
            for col_index, column_cells in enumerate(worksheet.columns, start=1):
                max_length = 0
                for cell in column_cells:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                worksheet.column_dimensions[get_column_letter(
                    col_index)].width = max_length + 2

            # --------------------------------------------------
            # SELECT A1 ON OPEN
            # --------------------------------------------------
            worksheet.sheet_view.selection = [
                Selection(activeCell="A1", sqref="A1")
            ]

        print(f"\nSaved Excel data file: {excel_file_name}")

    # --------------------------------------------------------------
    # HTML CONFIRMATION MESSAGE
    # --------------------------------------------------------------

    if INCLUDE_HTML:
        print(f"Saved combined HTML file: {html_file_name}")


# --------------------------------------------------------------
# ENTRY POINT
# --------------------------------------------------------------

if __name__ == "__main__":
    main()
