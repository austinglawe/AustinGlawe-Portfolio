import requests
from bs4 import BeautifulSoup
import pandas as pd
import json
import time

base_url = "https://www.lifetime.life"
locations_url = f"{base_url}/locations.html"
headers = {"User-Agent": "Mozilla/5.0"}

response = requests.get(locations_url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

clubs = []
current_state = None

# Build club list
for li in soup.find_all("li", class_="all-locations-item"):
    state_header = li.find_previous("h2", class_="h3")
    if state_header:
        current_state = state_header.get_text(strip=True)

    club_div = li.find("div", class_="small")
    link_tag = li.find("a", href=True)

    if club_div and link_tag:
        href = link_tag["href"]
        full_url = href if href.startswith("http") else base_url + href
        clubs.append({
            "State": current_state,
            "Club ID": club_div.get("id"),
            "Club Name": club_div.get_text(strip=True),
            "Link": full_url
        })

# Limit to first 5 for testing
# clubs = clubs[:5]

# Process each club
for club in clubs:
    try:
        resp = requests.get(club["Link"], headers=headers, timeout=10)
        sub_soup = BeautifulSoup(resp.text, "html.parser")

        json_script = sub_soup.find("script", type="application/ld+json")
        if json_script:
            data = json.loads(json_script.string)
            if isinstance(data, list):
                data = data[0]

            club["Phone"] = data.get("telephone", "")
            addr = data.get("address", {})
            club["Address"] = f"{addr.get('streetAddress', '')}, {addr.get('addressLocality', '')}, {addr.get('addressRegion', '')} {addr.get('postalCode', '')}"
            geo = data.get("geo", {})
            club["Latitude"] = geo.get("latitude", "")
            club["Longitude"] = geo.get("longitude", "")

            hours_by_day = {d: "" for d in [
                "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]}
            for item in data.get("openingHoursSpecification", []):
                day = item.get("dayOfWeek", "")
                opens = item.get("opens", "")
                closes = item.get("closes", "")
                if day:
                    hours_by_day[day] = f"{opens}–{closes}"
            for day, hrs in hours_by_day.items():
                club[f"{day} Hours"] = hrs

            features = data.get("amenityFeature", [])
            club["Amenity Features"] = "|".join(
                [f.get("value", "") for f in features if "value" in f])
        else:
            for field in ["Phone", "Address", "Latitude", "Longitude", "Amenity Features"]:
                club[field] = ""
            for day in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
                club[f"{day} Hours"] = ""

        # Membership info page
        membership_url = club["Link"].replace(".html", "/memberships.html")
        mem_resp = requests.get(membership_url, headers=headers, timeout=10)
        mem_soup = BeautifulSoup(mem_resp.text, "html.parser")

        # Updated: Loop through all tabs for membership prices
        club["Membership Tiers"] = ""
        club["Standard Price"] = ""
        club["Signature Price"] = ""
        club["26 and Under Price"] = ""
        club["65 Plus Price"] = ""

        for tab in mem_soup.select("div.tab-content > div.tab-pane"):
            membership_name = ""
            price = ""

            tab_id = tab.get("id", "")
            tab_button = mem_soup.select_one(
                f'button[data-bs-target="#{tab_id}"]')
            if tab_button:
                membership_name = tab_button.get_text(strip=True)

            price_span = tab.select_one(".join-offers-price .price")
            if price_span:
                price = price_span.get_text(strip=True)

            if "standard" in membership_name.lower():
                club["Standard Price"] = price
            elif "signature" in membership_name.lower():
                club["Signature Price"] = price
            elif "26" in membership_name:
                club["26 and Under Price"] = price
            elif "65" in membership_name:
                club["65 Plus Price"] = price

            club["Membership Tiers"] += f"{membership_name}: {price} | "

        def get_features_list_by_heading(heading_text):
            heading = mem_soup.find(
                lambda tag: tag.name == "h3" and heading_text in tag.text)
            if heading:
                ul = heading.find_next("ul")
                if ul:
                    return "|".join([li.get_text(strip=True) for li in ul.find_all("li")])
            return ""

        club["Exclusive Amenities"] = get_features_list_by_heading("Exclusive")
        club["Included in All"] = get_features_list_by_heading("Included")
        club["Not Included"] = get_features_list_by_heading("Not included")

    except Exception as e:
        print(f"Failed to process: {club['Link']} - {e}")
        default_fields = [
            "Phone", "Address", "Latitude", "Longitude", "Amenity Features",
            "Standard Price", "Signature Price", "26 and Under Price", "65 Plus Price",
            "Exclusive Amenities", "Not Included", "Included in All"
        ]
        for field in default_fields:
            club[field] = ""
        for day in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
            club[f"{day} Hours"] = ""

    time.sleep(0.5)

# Export to Excel
df = pd.DataFrame(clubs)
df.to_excel("lifetime_clubs_with_prices.xlsx", index=False)
print("Saved to 'lifetime_clubs_with_prices.xlsx'")
