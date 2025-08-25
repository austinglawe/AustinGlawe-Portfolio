values = [x * 2 for x in ranges(10)]
for x in values:
    print(x)


values = (x * 2 for x in ranges(10))
for x in values:
    print(x)

# generator objects are good for reducing memory size, however it won't give you the full number in that object
from sys import getsizeof

values = (x * 2 for x in ranges(100000))
print("gen:", getsizeof(values))

values = [x * 2 for x in ranges(100000)]
print("list:", getsizeof(values))











