from datetime import datetime
from meteostat import Hourly
import pandas as pd

# 1. Define the time span
start = datetime(2025, 1, 1)
end = datetime.now()   # <-- ensures you only request up-to-date data

# 2. Specify the two station IDs
stations = ["72698", "72793"]  # 72698 = PDX, 72793 = SEA

# 3. Fetch hourly data
data = Hourly(stations, start, end)
df = data.fetch()

# 4. (Optional) Rename station IDs to city names for clarity
mapping = {"72698": "Portland", "72793": "Seattle"}
df = df.reset_index().rename(columns={"station": "station_id"})
df["city"] = df["station_id"].map(mapping)

# 5. Re-order columns so “city” comes first
cols = ["city", "station_id", "time"] + \
    [c for c in df.columns if c not in ("city", "station_id", "time")]
df = df[cols]

# 6. Save to CSV
output_file = "portland_seattle_hourly_2025.csv"
df.to_csv(output_file, index=False)

print(f"Saved {len(df)} rows to {output_file}")
