class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y


point = Point(10, 20)
other = Point(1, 2)
print(point + other)



class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __add__(self, other):
        return Point(self.x + other.x, self.y + other.y)


point = Point(10, 20)
other = Point(1, 2)
combined = point + other
print(combined.x)
print(point + other)
















