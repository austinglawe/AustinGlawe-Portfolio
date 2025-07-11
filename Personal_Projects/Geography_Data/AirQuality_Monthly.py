import pandas as pd

# 1) Load the hourly CSV
df = pd.read_csv(
    "aq_open_meteo_all_vars_july1-10_2025.csv",
    parse_dates=["datetime"]
)

# 2) Index by timestamp
df = df.set_index("datetime")

# 3) Group by city, variable, latitude, longitude → resample monthly
monthly = (
    df
    .groupby(["city", "variable", "latitude", "longitude"])["value"]
    .resample("M")    # monthly bins
    .mean()           # average over each month
    .reset_index()
)

# 4) Rename columns for clarity
monthly = monthly.rename(columns={
    "datetime":  "month",
    "value":     "monthly_avg"
})

# 5) Save out
monthly.to_csv("aq_open_meteo_monthly_july1-10_2025.csv", index=False)
print(f"Saved {len(monthly)} monthly records to aq_open_meteo_monthly_july1-10_2025.csv")
