from meteostat import Stations
import pandas as pd

# 1. Fetch all stations
stations = Stations().fetch()

# 2. Select the metadata you want, including the built-in daily_start
df = (
    stations
    .reset_index()[[
        "id",           # station ID
        "name",         # station name
        "country",      # ISO country code
        "region",       # state / province code
        "latitude",
        "longitude",
        "daily_start"   # first day of daily observations
    ]]
    .rename(columns={
        "id":           "station_id",
        "name":         "station_name",
        "daily_start":  "first_daily"
    })
)

# 3. Reformat the date to YYYY.MM.DD
df["first_daily"] = df["first_daily"].dt.strftime("%Y.%m.%d")

# 4. Save to CSV
df.to_csv("global_stations_with_start.csv", index=False)
print(
    f"Written {len(df)} stations (with first_daily) to global_stations_with_start.csv")
