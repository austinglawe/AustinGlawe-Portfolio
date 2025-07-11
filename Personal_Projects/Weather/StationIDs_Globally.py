from meteostat import Stations
import pandas as pd

# 1. Query every station (no region filter)
stations = Stations().fetch()

# 2. Pick columns—including country
df = stations.reset_index()[[
    'id', 'name', 'country', 'region', 'latitude', 'longitude'
]].rename(columns={
    'id':        'station_id',
    'name':      'station_name',
    'country':   'country',      # ISO 3166-1 alpha-2 country code
    'region':    'region',       # state/province code (if applicable)
    'latitude':  'latitude',
    'longitude': 'longitude'
})

# 3. Save to CSV
df.to_csv('global_stations.csv', index=False)
print(f"Written {len(df)} global stations to global_stations.csv")
