import requests
from bs4 import BeautifulSoup

# Lifetime North Scottsdale Membership Page
url = "https://www.lifetime.life/locations/mn/new-hope/memberships.html"

# Define headers to mimic a real browser
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

# Send request
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

# Membership level name → corresponding content ID
membership_levels = {
    "Signature": "4b6d59ec-e69b-4f04-8085-904790b75eff",
    "Standard": "df08707b-b4b8-4c0c-ae42-e33a263a3a15",
    "26 & Under": "3e99115c-ef41-434f-8b37-5568e267d110",
    "65 Plus": "851efbb0-f572-4ca9-9672-a3a8f4b296f7"
}

# Loop through each membership level and extract pricing info
for level, div_id in membership_levels.items():
    target_block = soup.find("div", id=div_id)

    if target_block:
        price_block = target_block.find("div", class_="join-offers-price mb-3")
        if price_block:
            text = " ".join(price_block.stripped_strings)
            print(f"{level}: {text}")
        else:
            print(f"{level}: Pricing info not found.")
    else:
        print(f"{level}: Membership section not found.")
