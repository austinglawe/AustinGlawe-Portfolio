import pandas as pd

# 1) Read in your hourly file
df = pd.read_csv(
    "aq_open_meteo_all_vars_july1-10_2025.csv",
    parse_dates=["datetime"],
)

# 2) Set the datetime index
df = df.set_index("datetime")

# 3) Group by city, variable (and keep lat/lon) then resample daily
daily = (
    df
    .groupby(["city", "variable", "latitude", "longitude"])["value"]
    .resample("D")
    .mean()
    .reset_index()
)

# 4) Rename the date column
daily = daily.rename(columns={"datetime": "date", "value": "daily_avg"})

# 5) Save to CSV
daily.to_csv("aq_open_meteo_daily_july1-10_2025.csv", index=False)

print(f"Saved {len(daily)} daily records to aq_open_meteo_daily_july1-10_2025.csv")
