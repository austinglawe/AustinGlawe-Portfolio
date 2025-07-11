import requests
import pandas as pd

# ─── City coordinates ─────────────────────────────────────────────────────────
locations = {
    "Portland":    (45.5152,  -122.6784),
    "Seattle":     (47.6062,  -122.3321),
    "Minneapolis": (44.9778,   -93.2650),
    "San Diego":   (32.7157,  -117.1611)
}

# ─── Date range ───────────────────────────────────────────────────────────────
start_date = "2025-07-01"
end_date = "2025-07-10"

# ─── List every supported hourly variable ─────────────────────────────────────
hourly_vars = [
    "pm10", "pm2_5",
    "carbon_monoxide", "carbon_dioxide",
    "nitrogen_dioxide", "sulphur_dioxide",
    "ozone", "aerosol_optical_depth",
    "dust", "uv_index", "uv_index_clear_sky",
    "ammonia", "methane",
    "alder_pollen", "birch_pollen", "grass_pollen",
    "mugwort_pollen", "olive_pollen", "ragweed_pollen",
    "european_aqi", "european_aqi_pm2_5", "european_aqi_pm10",
    "european_aqi_nitrogen_dioxide", "european_aqi_ozone", "european_aqi_sulphur_dioxide",
    "us_aqi", "us_aqi_pm2_5", "us_aqi_pm10",
    "us_aqi_nitrogen_dioxide", "us_aqi_ozone", "us_aqi_sulphur_dioxide"
]
# Note: if you also want CH2O, PAN, VOCs, wildfire-PM, etc., add them here.

# ─── Fetch and flatten ────────────────────────────────────────────────────────
rows = []
for city, (lat, lon) in locations.items():
    params = {
        "latitude":   lat,
        "longitude":  lon,
        "start_date": start_date,
        "end_date":   end_date,
        "hourly":     ",".join(hourly_vars),
        "timezone":   "UTC"
    }
    r = requests.get(
        "https://air-quality-api.open-meteo.com/v1/air-quality",
        params=params
    )
    r.raise_for_status()
    js = r.json()

    times = js["hourly"]["time"]
    units = js.get("hourly_units", {})

    for var, values in js["hourly"].items():
        if var == "time":
            continue
        unit = units.get(var, "")
        for t, v in zip(times, values):
            rows.append({
                "city":      city,
                "datetime":  t,        # e.g. "2025-07-01T00:00"
                "variable":  var,
                "value":     v,
                "unit":      unit,
                "latitude":  lat,
                "longitude": lon
            })

# ─── Save to CSV ─────────────────────────────────────────────────────────────
df = pd.DataFrame(rows)
df.to_csv("aq_open_meteo_all_vars_july1-10_2025.csv", index=False)
print(
    f"Written {len(df)} rows for {len(locations)} cities and {len(hourly_vars)} variables.")
