print(type(5)) # <class 'int'>
print(type(range(5))) # <class 'range'>

# int is a primitive type, range is a complex type
# range is iterable - use it in a for loop
for x in range(5):
    print(x)

# strings are iterable
for x in "Python":
    print(x)

# Lists are iterable
for x in [1, 2, 3, 4]:
    print(x)

# later in the course:
for item in shopping_cart:
  print(item)


