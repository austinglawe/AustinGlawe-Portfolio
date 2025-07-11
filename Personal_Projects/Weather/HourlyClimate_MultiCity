from datetime import datetime
from meteostat import Hourly
import pandas as pd
import warnings

# Suppress missing-bulk warnings from Meteostat
warnings.filterwarnings("ignore", message="Cannot load hourly/")

# Define station IDs and their corresponding city names
stations = ["72698", "72793", "72658", "72290"]
mapping = {
    "72698": "Portland",
    "72793": "Seattle",
    "72658": "Minneapolis",
    "72290": "San Diego"
}

# Earliest available hourly data dates per station
inventory_start = {
    "72698": datetime(1936, 5, 1),
    "72793": datetime(1948, 1, 1),
    "72658": datetime(1945, 1, 1),
    "72290": datetime(1942, 1, 1)
}

# Overall time range
start = min(inventory_start.values())
end = datetime.now()

# CSV splitting parameters
MAX_ROWS = 1_000_000
file_index = 1
rows_written = 0
output_prefix = "hourly_data"
current_file = f"{output_prefix}_{file_index}.csv"
header_written = False

# Loop through each station, fetch hourly data, and write to CSV chunks
for station_id in stations:
    df = Hourly(station_id, inventory_start[station_id], end).fetch()
    if df.empty:
        continue

    # Reset index to turn the timestamp into a 'time' column
    df = df.reset_index()

    # Add station metadata
    df["station_id"] = station_id
    df["city"] = mapping[station_id]

    # Reorder columns: city, station_id, time, then all weather fields
    weather_cols = [c for c in df.columns if c not in (
        "city", "station_id", "time")]
    df = df[["city", "station_id", "time"] + weather_cols]

    # Write in 1,000,000‐row chunks
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

        # If max rows reached, start a new file
        if rows_written >= MAX_ROWS:
            file_index += 1
            current_file = f"{output_prefix}_{file_index}.csv"
            rows_written = 0
            header_written = False

print(f"Done! CSV files created with prefix '{output_prefix}_<n>.csv'.")
