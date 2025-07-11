from datetime import datetime
import pandas as pd
import warnings
from meteostat import Hourly

# ─── Suppress Meteostat bulk‐file warnings ─────────────────────────────────────
warnings.filterwarnings("ignore", message="Cannot load hourly/")

# ─── Time span ────────────────────────────────────────────────────────────────
start = datetime(2025, 7, 10)
end = datetime.now()

# ─── Load station metadata ────────────────────────────────────────────────────
stations_df = pd.read_csv('global_stations.csv', dtype={'station_id': str})

# Build mappings for attaching metadata
mappings = {}
for col in ['station_name', 'country', 'region', 'latitude', 'longitude']:
    if col in stations_df.columns:
        mappings[col] = dict(zip(stations_df['station_id'], stations_df[col]))

# ─── CSV Splitting Parameters ─────────────────────────────────────────────────
MAX_ROWS = 1_000_000
file_index = 1
rows_written = 0
output_prefix = 'hourly_data_2025'
current_file = f"{output_prefix}_{file_index}.csv"
header_written = False

# ─── Loop through every station ───────────────────────────────────────────────
for station_id in stations_df['station_id']:
    # Fetch hourly observations
    df = Hourly(station_id, start, end).fetch()
    if df.empty:
        continue

    # Turn the index into a column; add station_id
    df = df.reset_index()
    df['station_id'] = station_id

    # Attach metadata (name, country, region, lat, lon)
    for col, mapping in mappings.items():
        df[col] = mapping.get(station_id)

    # Reorder so metadata comes first, then time, then weather fields
    meta_cols = ['station_id'] + list(mappings.keys())
    weather_cols = [c for c in df.columns if c not in meta_cols + ['time']]
    df = df[meta_cols + ['time'] + weather_cols]

    # Write out in 1,000,000-row chunks
    start_loc = 0
    total_rows = len(df)
    while start_loc < total_rows:
        available = MAX_ROWS - rows_written
        end_loc = min(start_loc + available, total_rows)
        chunk = df.iloc[start_loc:end_loc]

        chunk.to_csv(
            current_file,
            mode='a',
            header=(not header_written),
            index=False
        )

        header_written = True
        rows_written += len(chunk)
        start_loc = end_loc

        # Rotate to a new file if needed
        if rows_written >= MAX_ROWS:
            file_index += 1
            current_file = f"{output_prefix}_{file_index}.csv"
            rows_written = 0
            header_written = False

print(f"Done: files written with prefix `{output_prefix}_<n>.csv`.")
