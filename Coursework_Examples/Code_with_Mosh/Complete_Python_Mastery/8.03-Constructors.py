class Point:
    def __init__(self, x, y):    # magic method called instructor - underlined
        self.x = x
        self.y = y
      
    def draw(self):
        print(f"Point ({self.x}, {self.y})")


point = Point(1,2)
print(point.x)
point.draw()
















