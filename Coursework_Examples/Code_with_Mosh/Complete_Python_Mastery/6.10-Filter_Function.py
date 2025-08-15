# We have a list of items.. We want to filter this list, and only return items with price greater than or equal to 10

items = [
  ("Product1", 10),
  ("Product2", 9),
  ("Product3", 12),
]

x = filter(lambda item: item[1] >= 10, items)
print(x) # <filter object at 0x105bc23c8>

items = [
  ("Product1", 10),
  ("Product2", 9),
  ("Product3", 12),
]

filtered_list = list(filter(lambda item: item[1] >= 10, items))
print(filtered_list)










