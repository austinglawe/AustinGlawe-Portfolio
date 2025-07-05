import pandas as pd

url = "https://en.wikipedia.org/wiki/List_of_national_historic_sites_and_historical_parks_of_the_United_States"

# 1) Read every table on the page
all_tables = pd.read_html(url)

# 2) Find the NHS table by checking for its key columns
nhs_tables = [
    tbl for tbl in all_tables
    if {"Name", "Location", "Description"}.issubset(tbl.columns)
]

if not nhs_tables:
    raise ValueError(
        "Could not find the National Historic Sites table on Wikipedia.")

# 3) We expect exactly one match – use it
nhs = nhs_tables[0]

# should be 86 as of mid-2025
print(f"Found {len(nhs)} National Historic Sites.")

# 4) Export to CSV
nhs.to_csv("national_historic_sites.csv", index=False)
