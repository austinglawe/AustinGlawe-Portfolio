from meteostat import Stations
import pandas as pd

# 1. Fetch all station metadata
stations = Stations().fetch()

# 2. Reset index and select every column you can
df = stations.reset_index()[[
    "id",            # station ID
    "name",          # station name
    "country",       # country code
    "region",        # state/province code
    "latitude",
    "longitude",
    "elevation",     # meters above sea level
    "timezone",      # IANA time-zone string
    "wmo",           # WMO station ID
    "icao",          # ICAO code
    "hourly_start",  # first day of hourly data
    "hourly_end",    # last day of hourly data
    "daily_start",   # first day of daily data
    "daily_end",     # last day of daily data
    "monthly_start", # first day of monthly data
    "monthly_end"    # last day of monthly data
]]

# 3. Rename to more descriptive field names
df = df.rename(columns={
    "id":            "station_id",
    "name":          "station_name",
    "region":        "state_or_region",
    "wmo":           "wmo_id",
    "icao":          "icao_id",
    "hourly_start":  "first_hourly",
    "hourly_end":    "last_hourly",
    "daily_start":   "first_daily",
    "daily_end":     "last_daily",
    "monthly_start": "first_monthly",
    "monthly_end":   "last_monthly"
})

# 4. Format all of the *_hourly/_daily/_monthly columns as YYYY.MM.DD
date_cols = [
    "first_hourly","last_hourly",
    "first_daily","last_daily",
    "first_monthly","last_monthly"
]

df[date_cols] = df[date_cols].apply(lambda col: col.dt.strftime("%Y.%m.%d"))

# 5. Save to CSV
df.to_csv("global_stations_full_metadata.csv", index=False)
print(f"Written {len(df)} stations with full metadata to global_stations_full_metadata.csv")
