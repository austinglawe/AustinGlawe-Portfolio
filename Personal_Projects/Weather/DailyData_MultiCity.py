from datetime import datetime
from meteostat import Daily
import pandas as pd
import warnings

# ─── Suppress missing-bulk warnings ──────────────────────────────────────────
warnings.filterwarnings("ignore", message="Cannot load daily/")

# ─── Stations & city names ───────────────────────────────────────────────────
stations = ["72698", "72793", "72658", "72290"]
mapping = {
    "72698": "Portland",
    "72793": "Seattle",
    "72658": "Minneapolis",
    "72290": "San Diego"
}

# ─── Station-specific first‐available dates ───────────────────────────────────
inventory_start = {
    "72698": datetime(1936, 5, 1),
    "72793": datetime(1948, 1, 1),
    "72658": datetime(1945, 1, 1),
    "72290": datetime(1942, 1, 1)
}

start = min(inventory_start.values())
end   = datetime.now()

# ─── CSV splitting parameters ─────────────────────────────────────────────────
MAX_ROWS       = 1_000_000
file_index     = 1
rows_written   = 0
output_prefix  = "daily_data"
current_file   = f"{output_prefix}_{file_index}.csv"
header_written = False

# ─── Fetch & write loop ───────────────────────────────────────────────────────
for station_id in stations:
    # Pull daily data
    df = Daily(station_id, inventory_start[station_id], end).fetch()
    if df.empty:
        continue

    # Turn the index into a 'time' column
    df = df.reset_index()

    # Add station metadata
    df["station_id"] = station_id
    df["city"]       = mapping[station_id]

    # Re‐order so: city, station_id, time, then all daily fields
    daily_cols = [c for c in df.columns if c not in ("city", "station_id", "time")]
    df = df[["city", "station_id", "time"] + daily_cols]

    # Write out in 1 000 000-row chunks
    start_loc  = 0
    total_rows = len(df)
    while start_loc < total_rows:
        space_left = MAX_ROWS - rows_written
        end_loc    = min(start_loc + space_left, total_rows)
        chunk      = df.iloc[start_loc:end_loc]

        chunk.to_csv(
            current_file,
            mode="a",
            header=not header_written,
            index=False
        )

        header_written = True
        rows_written  += len(chunk)
        start_loc      = end_loc

        # Rotate to next file if needed
        if rows_written >= MAX_ROWS:
            file_index     += 1
            current_file    = f"{output_prefix}_{file_index}.csv"
            rows_written    = 0
            header_written  = False

print(f"Done! CSVs created as {output_prefix}_1.csv, {output_prefix}_2.csv, …")
