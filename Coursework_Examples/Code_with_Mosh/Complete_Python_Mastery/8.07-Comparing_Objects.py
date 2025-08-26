class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

point = Point(1, 2)
other = Point(1, 2)
print(point == other)
# referencing 2 different pieces of 



class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __eq__(self, other):
        return self.x == other.x and self.y == other.y

point = Point(1, 2)
other = Point(1, 2)
print(point == other)
print(point > other) # TypeError: '>' not supported between instances of 'Point' and 'Point'



class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __eq__(self, other):
        return self.x == other.x and self.y == other.y

    def __gt__(self, other):
        return self.x > other.x and self.y > other.y

point = Point(10, 20)
other = Point(1, 2)
print(point > other)
print(point < other)












