# -----------------------------------------
# Python_File_IO:
# Working with CSV and JSON files
# -----------------------------------------
#
# Reading CSV:
#   import csv
#   with open("data.csv", newline='') as csvfile:
#       reader = csv.reader(csvfile)
#       for row in reader:
#           print(row)
#
# Writing CSV:
#   with open("output.csv", "w", newline='') as csvfile:
#       writer = csv.writer(csvfile)
#       writer.writerow(["Name", "Age", "City"])
#       writer.writerow(["Austin", 30, "Minneapolis"])
#
# Reading JSON:
#   import json
#   with open("data.json", "r") as jsonfile:
#       data = json.load(jsonfile)
#       print(data)
#
# Writing JSON:
#   data = {"name": "Austin", "age": 30, "city": "Minneapolis"}
#   with open("output.json", "w") as jsonfile:
#       json.dump(data, jsonfile, indent=4)
#
# Best practices:
# - Use newline='' for CSV files.
# - Use indent=4 for readable JSON.
# - Handle exceptions for file and JSON errors.
#
# -----------------------------------------
