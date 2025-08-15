list1 = [1, 2, 3]
list2 = [10, 20, 30]

# We want a result like this: [(1, 10), (2, 20), (3, 30)]

print(zip(list1, list2)) # <zip object at 0x10fee2348>

print(list(zip(list1, list2)))


