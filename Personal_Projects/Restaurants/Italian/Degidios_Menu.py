import requests
from bs4 import BeautifulSoup
import pandas as pd

# Step 1: Fetch the webpage
url = "https://degidios.com/menu"
headers = {"User-Agent": "Mozilla/5.0"}

response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.text, 'html.parser')

# Step 2: Initialize
menu_data = []
seen = set()
current_category = None

# Step 3: Traverse through relevant elements
for element in soup.find_all(["h3", "div", "p"]):

    # Update current category
    if element.name == "h3" and "framer-styles-preset-b7s48y" in element.get("class", []):
        current_category = element.get_text(strip=True)

    # Process each menu item block
    elif element.name == "div" and "framer-ccp2t7" in element.get("class", []):
        # Extract name
        name_tag = element.find("p", class_="framer-styles-preset-1d2so12")
        name = name_tag.get_text(strip=True) if name_tag else ""

        # Extract price
        price_tag = element.find("p", class_="framer-styles-preset-p5viac")
        price = price_tag.get_text(strip=True) if price_tag else ""

        # Extract description placeholder (will populate in the next block if found)
        description = ""

        # Extract tags: GF, V, etc., deduplicated and uppercased
        tag_list = []
        for tag in element.find_all("p", class_="framer-text"):
            tag_text = tag.get_text(strip=True).upper()
            if tag_text in ["GF", "V", "VG", "DF"] and tag_text not in tag_list:
                tag_list.append(tag_text)

        tags = " | ".join(tag_list)

        # Skip duplicates
        key = (name, price, description, tags, current_category)
        if key in seen:
            continue
        seen.add(key)

        # Store the cleaned row
        menu_data.append({
            "Category": current_category,
            "Name": name,
            "Description": description,  # will be updated below
            "Price": price,
            "Tags": tags
        })

    # Improved description logic: match any class that contains '1bg58z'
    elif element.name == "p" and any("1bg58z" in cls for cls in element.get("class", [])):
        if menu_data:
            menu_data[-1]["Description"] = element.get_text(strip=True)

# Step 4: Save to Excel
df = pd.DataFrame(menu_data)
output_file = "degidios_menu_clean.xlsx"
df.to_excel(output_file, index=False)
print(f"Saved {len(df)} unique menu items to '{output_file}'")
