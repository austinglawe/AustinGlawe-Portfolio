numbers = [1, 2, 3]
first = numbers[0]
second = numbers[1]
third = numbers[2]

# cleaner way: list unpacking like below
numbers = [1, 2, 3]
first, second, third = numbers
# important: the number of variables on the left side, should equal the number on the right side of the equal sign

first, second = numbers # this will give an error

# if you only want a couple, you can do the following
numbers = [1, 2, 3, 4, 4, 4, 4, 4, 4]
first, second, *other = numbers
print(first)
print(second)
print(other)


numbers = [1, 2, 3, 4, 4, 4, 4, 4, 4]
first, *other, last = numbers
print(first)
print(other)
print(last)









