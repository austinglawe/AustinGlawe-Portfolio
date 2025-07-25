import requests
from bs4 import BeautifulSoup
import pandas as pd
import json
import time
from urllib.parse import urljoin

base_url = "https://www.lifetime.life"
locations_url = f"{base_url}/locations.html"
headers = {"User-Agent": "Mozilla/5.0"}

response = requests.get(locations_url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

clubs = []
current_state = None

# Step 1: Gather club links and metadata
for li in soup.find_all("li", class_="all-locations-item"):
    state_header = li.find_previous("h2", class_="h3")
    if state_header:
        current_state = state_header.get_text(strip=True)

    club_div = li.find("div", class_="small")
    link_tag = li.find("a", href=True)

    if club_div and link_tag:
        clubs.append({
            "State": current_state,
            "Club ID": club_div.get("id"),
            "Club Name": club_div.get_text(strip=True),
            "Link": urljoin(base_url, link_tag["href"])  # Ensure proper URL
        })

# Limit to first 5 for testing
# clubs = clubs[:5]

# Step 2: Crawl each club's JSON-LD data
for club in clubs:
    try:
        resp = requests.get(club["Link"], headers=headers, timeout=10)
        sub_soup = BeautifulSoup(resp.text, "html.parser")

        json_script = sub_soup.find("script", type="application/ld+json")
        if json_script:
            data = json.loads(json_script.string)[0] if json_script.string.strip(
            ).startswith("[") else json.loads(json_script.string)

            club["Phone"] = data.get("telephone", "")
            addr = data.get("address", {})
            club["Address"] = f"{addr.get('streetAddress', '')}, {addr.get('addressLocality', '')}, {addr.get('addressRegion', '')} {addr.get('postalCode', '')}"
            geo = data.get("geo", {})
            club["Latitude"] = geo.get("latitude", "")
            club["Longitude"] = geo.get("longitude", "")

            # Amenity Features
            amenities = data.get("AmenityFeature", [])
            club["Amenities"] = " | ".join(
                [item.get("value", "") for item in amenities])

            # Weekly hours
            hours_by_day = {d: "" for d in [
                "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]}
            for item in data.get("openingHoursSpecification", []):
                day = item.get("dayOfWeek", "")
                opens = item.get("opens", "")
                closes = item.get("closes", "")
                if day:
                    hours_by_day[day] = f"{opens}–{closes}"

            club["Today's Hours"] = ""
            for day, hrs in hours_by_day.items():
                club[f"{day} Hours"] = hrs

        else:
            club.update({k: "" for k in [
                        "Phone", "Address", "Latitude", "Longitude", "Today's Hours", "Amenities"]})
            for day in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
                club[f"{day} Hours"] = ""

    except Exception as e:
        print(f"Failed to process: {club['Link']} - {e}")
        club.update({k: "" for k in [
                    "Phone", "Address", "Latitude", "Longitude", "Today's Hours", "Amenities"]})
        for day in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
            club[f"{day} Hours"] = ""

    time.sleep(0.5)

# Step 3: Export to Excel
df = pd.DataFrame(clubs)
df.to_excel("lifetime_clubs_json_ld.xlsx", index=False)
print("Saved to 'lifetime_clubs_json_ld.xlsx'")

