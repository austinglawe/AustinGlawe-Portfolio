def multiply(x, y):
    return x * y

multiply(2, 3)

# if you wanted to do this:
# multiply(2, 3, 4, 5)

# use a plural parameter to show a collection of arguments:
# turns it into a tuple
def multiply(*numbers):
    for number in numbers:
        print(number)

multiply(2, 3, 4, 5)


def multiply(*numbers):
    total = 1
    for number in numbers:
        total *= number
    return total

print(multiply(2, 3, 4, 5))
