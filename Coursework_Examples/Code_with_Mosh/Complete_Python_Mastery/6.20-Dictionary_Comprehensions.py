values = []
for x in range(5):
    values.append(x * 2)

# [expression for item in items]

values = [x * 2 for x in range(5)]
print(values)


values = {x * 2 for x in range(5)}
print(values)

# {1, 2, 3, 4} # Set
# {1: "a", 2: "b"} # Dictionary

# this is the same:
# values = {x: x * 2 for x in range(5)}
# as:
# values = {}
# for x in range(5):
    # values[x] = x * 2


# tuples
values = (x * 2 for x in range(5))
print(values) # <generator object <genexpr> at 0x109c13cf0>















