import requests
import pandas as pd
from datetime import datetime, timedelta

# Example: get sunrise/sunset for a list of cities and dates
locations = {
    "Portland": (45.5152, -122.6784),
    "Seattle":  (47.6062, -122.3321),
    "Minneapolis": (44.9778, -93.2650),
    "San Diego": (32.7157, -117.1611)
}

# Build a DataFrame of every date in 2025 for each city
dates = pd.date_range("2025-01-01", datetime.now().date(), freq="D")
rows = []
for city, (lat, lon) in locations.items():
    for date in dates:
        r = requests.get(
            "https://api.sunrise-sunset.org/json",
            params={
                "lat": lat,
                "lng": lon,
                "date": date.strftime("%Y-%m-%d"),
                "formatted": 0
            }
        ).json()["results"]
        rows.append({
            "city": city,
            "date": date,
            **r
        })

df = pd.DataFrame(rows)
df.to_csv("sunrise_sunset_2025.csv", index=False)
