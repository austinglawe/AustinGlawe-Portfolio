x = 10
y = 11

# need a third variable
z = x
x = y
y = z

print("x", x)
print("y", y)

# in python you can actually do this without a third variable
x, y = y, x
# essentially it is like tuple unpacking


