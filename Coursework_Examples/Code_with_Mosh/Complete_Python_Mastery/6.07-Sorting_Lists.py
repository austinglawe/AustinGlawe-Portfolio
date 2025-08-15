numbers = [3, 51, 2, 8, 6]
numbers.sort() # Default - Ascending order
print(numbers)

# Descending order
numbers.sort(reverse=True)
print(numbers)

print(sorted(numbers))
print(numbers)



items = [
  ("Product1", 10),
  ("Product2", 9),
  ("Product3", 12),
]
items.sort
print(items) # not sorted.

def sort_item(item):
    return item[1]

items.sort(sort_item) # TypeError: sort() takes no positional arguments
print(items)

# instead do this:
items.sort(key=sort_item)
print(items)









