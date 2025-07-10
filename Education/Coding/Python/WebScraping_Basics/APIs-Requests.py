# -----------------------------------------
# Python_APIs:
# Consuming REST APIs with requests
# -----------------------------------------
#
# GET request:
#   import requests
#   response = requests.get("https://api.example.com/data")
#   if response.status_code == 200:
#       data = response.json()
#
# POST request with JSON:
#   payload = {"name": "Austin", "age": 30}
#   response = requests.post("https://api.example.com/users", json=payload)
#
# Headers and authentication:
#   headers = {"Authorization": "Bearer YOUR_TOKEN"}
#   response = requests.get("https://api.example.com/protected", headers=headers)
#
# Best practices:
# - Check response.status_code
# - Handle exceptions (timeouts, errors)
# - Use HTTPS
# - Respect rate limits
# - Store tokens securely
#
# -----------------------------------------
