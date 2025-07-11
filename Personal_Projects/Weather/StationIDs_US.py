from meteostat import Stations
import pandas as pd

# 1. Query all U.S. stations
stations = Stations().region('US').fetch()

# 2. Select and rename the fields you need
df = stations.reset_index()[[
    'id', 'name', 'region', 'latitude', 'longitude'
]].rename(columns={
    'id':        'station_id',
    'name':      'station_name',
    'region':    'state',
    'latitude':  'latitude',
    'longitude': 'longitude'
})

# 3. Save to CSV
df.to_csv('us_stations_with_coords.csv', index=False)
print(f"Written {len(df)} U.S. stations across all states to us_stations_with_coords.csv")
