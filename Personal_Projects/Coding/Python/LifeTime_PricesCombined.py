import requests
from bs4 import BeautifulSoup

# Target URL
url = "https://www.lifetime.life/locations/az/north-scottsdale/memberships.html"

# Headers to mimic a real browser
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

# Send the request
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

# Membership levels and their corresponding IDs
membership_levels = {
    "Signature": "4b6d59ec-e69b-4f04-8085-904790b75eff",
    "Standard": "df08707b-b4b8-4c0c-ae42-e33a263a3a15",
    "26 & Under": "3e99115c-ef41-434f-8b37-5568e267d110",
    "65 Plus": "851efbb0-f572-4ca9-9672-a3a8f4b296f7"
}

# Dictionary to store the price for each level
prices = {}

# Loop through and extract prices
for level, div_id in membership_levels.items():
    target_block = soup.find("div", id=div_id)

    if target_block:
        price_span = target_block.find("span", class_="h2 price")
        price = price_span.get_text(strip=True) if price_span else "N/A"
    else:
        price = ""

    prices[level] = price

# Combine into a single string separated by "|"
price_string = " | ".join(
    [f"{level}: {price}" for level, price in prices.items()])
print(price_string)
