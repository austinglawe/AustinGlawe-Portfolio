# -----------------------------------------
# Python_Data_Analysis:
# Using pandas and NumPy
# -----------------------------------------
#
# NumPy basics:
#   import numpy as np
#   arr = np.array([1, 2, 3, 4])
#   print(arr.mean())
#   print(arr + 10)
#
# pandas basics:
#   import pandas as pd
#   data = {"Name": ["Austin", "Emma", "Liam"], "Age": [30, 25, 28], "Score": [85, 92, 88]}
#   df = pd.DataFrame(data)
#   print(df.head())
#   print(df["Age"].mean())
#
# Read CSV:
#   df = pd.read_csv("data.csv")
#
# DataFrame operations:
#   filtered = df[df["Age"] > 25]
#   df["Passed"] = df["Score"] > 80
#   grouped = df.groupby("Passed")["Score"].mean()
#
# Best practices:
# - Use pandas for tabular data.
# - Use NumPy for numeric arrays.
# - Chain operations clearly.
# - Inspect data regularly.
#
# -----------------------------------------
