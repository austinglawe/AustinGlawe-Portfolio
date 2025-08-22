# a collection with no duplicates
numbers = [1, 1, 2, 3, 4]
uniques = set(numbers)
print(uniques)
second = {1, 4}
second.add(5)
# notice sets use curly braces
second.remove(5)
len(second)

numbers = [1, 1, 2, 3, 4]
first = set(numbers)
second = {1, 5}

# power comes from
# use the pipe to combine sets

print(first | second)
# union of sets where if the number is in one or the other, it will be returned

print(first & second)
# returns what is in BOTH first and second sets

print(first - second)
# returns the values in the first that are not in the second

print(first ^ second)
# returns what is in one of the two sets but not in both

print(first[0]) # TypeError: 'set' object does not support indexing
# instead

if 1 in first:
    print("yes")





