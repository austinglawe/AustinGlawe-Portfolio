items = [
  ("Product1", 10),
  ("Product2", 9),
  ("Product3", 12),
]

# List comprehensions are cleaner to read than the map and filter functions:

prices = list(map(lambda item: items[1], items))
# List comprension is : [expression for item in items]
prices = [item[1] for item in items]

filtered = list(filter(lambda item: item[1] >=10, items)
# filtered = [expression for item in items]
filtered = [item for item in items if item[1] >= 10]



                
