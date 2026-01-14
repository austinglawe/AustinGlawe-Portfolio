"""
Step 2:
    - Fetch a small sample of stations
    - Keep only the columns we care about right now
    - Add a 'data_source' column
    - Just print the result
"""

from meteostat import Stations

# 1. Fetch a small sample
stations_df = Stations().fetch(5)

print("Original columns:")
print(stations_df.columns)

# 2. Select and rename columns to match our plan
#    Meteostat uses 'elevation', so we'll rename it to 'elevation_m'
stations_selected = stations_df.reset_index().rename(
    columns={
        "id": "station_id",       # index becomes 'id' after reset_index()
        "name": "name",
        "country": "country",
        "region": "region",
        "latitude": "latitude",
        "longitude": "longitude",
        "elevation": "elevation_m",
    }
)

# 3. Keep only the columns we care about so far
stations_selected = stations_selected[
    ["station_id", "name", "country", "region",
     "latitude", "longitude", "elevation_m"]
]

# 4. Add a data_source column with a constant value
stations_selected["data_source"] = "meteostat"

print("\nShaped stations table:")
print(stations_selected)
