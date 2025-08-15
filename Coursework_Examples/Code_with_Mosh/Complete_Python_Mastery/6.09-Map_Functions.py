items = [
  ("Product1", 10),
  ("Product2", 9),
  ("Product3", 12),
]


prices = []
for item in items:
    prices.append(item[1])

print(prices)

# instead of the loop we can do this:
items = [
  ("Product1", 10),
  ("Product2", 9),
  ("Product3", 12),
]

x = map(lambda item: item[1],items)
print(x) # <map object at 0x108c4f3c8>


items = [
  ("Product1", 10),
  ("Product2", 9),
  ("Product3", 12),
]


x = map(lambda item: item[1],items)
for item in x:
    print(item)


# you can also make it into a list

items = [
  ("Product1", 10),
  ("Product2", 9),
  ("Product3", 12),
]


prices = list(map(lambda item: item[1],items))
print(prices)





