class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y
      
    def draw(self):
        print(f"Point ({self.x}, {self.y})")

point = Point(1, 2)
point.draw()


point = Point(0, 0)
point = Point.zero() # factory method - creates a new method
point.draw()


class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y
      
    @classmethod # decorator
    def zero(cls):
        return cls(0, 0)
  
    def draw(self):
        print(f"Point ({self.x}, {self.y})")

point = Point.zero()
point.draw()










