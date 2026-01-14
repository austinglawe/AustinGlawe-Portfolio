# First we need to install the libraries: "meteostat" AND "pandas"
# pip install meteostat pandas

"""
Step 1:
    - Ask Meteostat for a small sample of weather stations
    - Look at the result in Python (no saving yet)
"""

from meteostat import Stations

# 1. Create a Stations object (this is like starting a query)
stations_query = Stations()

# 2. Fetch a SMALL number of stations as a table (DataFrame)
#    Here we ask for just 5 rows so it is easy to read.
stations_df = stations_query.fetch(5)

# 3. Print what type of object this is
print("Type of stations_df:", type(stations_df))

# 4. Print the whole small table
print("\nStations sample:")
print(stations_df)
