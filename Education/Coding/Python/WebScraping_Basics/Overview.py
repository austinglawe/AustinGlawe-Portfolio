# -----------------------------------------
# Python_Web_Scraping:
# Basics with requests, BeautifulSoup, Selenium
# -----------------------------------------
#
# Download page with requests:
#   import requests
#   response = requests.get("https://example.com")
#   if response.status_code == 200:
#       html = response.text
#
# Parse HTML with BeautifulSoup:
#   from bs4 import BeautifulSoup
#   soup = BeautifulSoup(html, "html.parser")
#   title = soup.title.string
#
# Find elements:
#   headings = soup.find_all("h2")
#   for h in headings:
#       print(h.text)
#
# Dynamic content with Selenium:
#   from selenium import webdriver
#   driver = webdriver.Chrome()
#   driver.get("https://example.com")
#   content = driver.page_source
#   driver.quit()
#
# Best practices:
# - Respect robots.txt and terms.
# - Avoid too many rapid requests.
# - Use user-agent headers.
# - Handle errors gracefully.
#
# -----------------------------------------
