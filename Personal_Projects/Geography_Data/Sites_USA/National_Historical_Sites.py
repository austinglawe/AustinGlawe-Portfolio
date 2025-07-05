import pandas as pd

url = "https://en.wikipedia.org/wiki/List_of_national_historic_sites_and_historical_parks_of_the_United_States"
tables = pd.read_html(url)
nhs = tables[1]   # table[1] is the “National Historic Sites” table
nhs.to_csv("national_historic_sites.csv", index=False)
