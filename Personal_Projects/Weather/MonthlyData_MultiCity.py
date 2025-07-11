from datetime import datetime
from meteostat import Monthly
import pandas as pd
import warnings

# Suppress missing‐bulk warnings for monthly data
warnings.filterwarnings("ignore", message="Cannot load monthly/")

# Define the station IDs and their friendly names
stations = ["72698", "72793", "72658", "72290"]
mapping = {
    "72698": "Portland",
    "72793": "Seattle",
    "72658": "Minneapolis",
    "72290": "San Diego"
}

# Earliest available monthly‐data dates per station
inventory_start = {
    "72698": datetime(1936, 5, 1),
    "72793": datetime(1948, 1, 1),
    "72658": datetime(1945, 1, 1),
    "72290": datetime(1942, 1, 1)
}

# Overall time span: from the earliest station start to now
start = min(inventory_start.values())
end = datetime.now()

# CSV splitting settings
MAX_ROWS = 1_000_000
file_index = 1
rows_written = 0
output_prefix = "monthly_data"
current_file = f"{output_prefix}_{file_index}.csv"
header_written = False

# Loop through each station, fetch monthly data, and write in chunks
for station_id in stations:
    df = Monthly(station_id, inventory_start[station_id], end).fetch()
    if df.empty:
        continue

    # Reset index to turn the timestamp into a 'time' column
    df = df.reset_index()

    # Add station metadata
    df["station_id"] = station_id
    df["city"] = mapping[station_id]

    # Reorder so: city, station_id, time, then all monthly fields
    monthly_cols = [c for c in df.columns if c not in (
        "city", "station_id", "time")]
    df = df[["city", "station_id", "time"] + monthly_cols]

    # Write out in 1,000,000-row chunks
    start_loc = 0
    total_rows = len(df)
    while start_loc < total_rows:
        space_left = MAX_ROWS - rows_written
        end_loc = min(start_loc + space_left, total_rows)
        chunk = df.iloc[start_loc:end_loc]

        chunk.to_csv(
            current_file,
            mode="a",
            header=not header_written,
            index=False
        )

        header_written = True
        rows_written += len(chunk)
        start_loc = end_loc

        # Roll over to the next file when limit is reached
        if rows_written >= MAX_ROWS:
            file_index += 1
            current_file = f"{output_prefix}_{file_index}.csv"
            rows_written = 0
            header_written = False

print(f"Done! CSV files created with prefix '{output_prefix}_<n>.csv'.")
