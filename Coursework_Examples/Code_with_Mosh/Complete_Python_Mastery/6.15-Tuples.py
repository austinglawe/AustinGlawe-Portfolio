# Tuples are read only.
# we cannot remove existing objects, we cannot modify it, we cannot add items to it.
point = (1, 2)

# you can also exclude the parentheses
point = 1, 2
print(type(point))

# but if you only have 1 object, add a trailing comma
point = 1
print(type(point)) # <clas 'int'>
point = 1,
print(type(point)) # <class 'tuple'>

# to define an empty tuple use empty parentheses
point = ()
print(type(point)) # <class 'tuple'>


point = (1, 2) + (3, 4)
print(point)

point = (1, 2) * 3
print(point)

point = tuple([1, 2])
print(point)

point = tuple("Hello World")
print(point)

point = (1, 2, 3)
print(point[0:2])

x, y, z = point
if 10 in point:
    print("exists")

point[0] = 10 # TypeError: 'tuple' object does not support item assignment















