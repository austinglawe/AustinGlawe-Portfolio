# collection of key:value pairs
# map a key to a value

### phone book in real life
point = {"x": 1, "y": 2}

# list()
# tuple()
# set()
# dict()

point = dict(x=1, y=2)
# index of a dictionary is the name of the key
print(point["x"])
point["x"] = 10
print(point)

point["z"] = 20
print(point)
print(point["a"]) # KeyError: 'a'

# a workaround
if "a" in point:
    print(point["a"])

print(point.get("a")) # None

print(point.get("a", 0)) # 0 - you can decide what it returns by default if it is not found

# delete an item
del point["x"]
print(point)

for key in point:
    print(key, point[key])








