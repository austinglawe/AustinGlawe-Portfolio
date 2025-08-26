class Point:
    def __init__(self, x, y):
        self.x = x # instance attributes
        self.y = y
      
    def draw(self):
        print(f"Point ({self.x}, {self.y})")


point = Point(1,2)
point.z = 10
point.draw()


another = Point(3, 4)
another.draw()



class Point:
    default_color = "red" # class attribute - applies to all methods in this class
  
    def __init__(self, x, y):
        self.x = x
        self.y = y
      
    def draw(self):
        print(f"Point ({self.x}, {self.y})")

point = Point(1, 2)
print(point.default_color)
print(Point.default_color)
point.draw()

Point.default_color = "yellow"
print(point.default_color)
print(Point.default_color)









