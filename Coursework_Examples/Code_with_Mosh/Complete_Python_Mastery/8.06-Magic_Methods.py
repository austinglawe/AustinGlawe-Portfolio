# https://rszalski.github.io/magicmethods/#representations
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def draw(self):
        print(f"Point ({self.x}, {self.y})")


point = Point(1, 2)
print(point) # <__main__.Point object at 0x107330208>


class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __str__(self):
        return f"({self.x}, {self.y})")
  
    def draw(self):
        print(f"Point ({self.x}, {self.y})")

point = Point(1, 2)
print(str(point))














