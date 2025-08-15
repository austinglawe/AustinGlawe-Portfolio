def greet(name):
    print(f"Hi {name}")


# There are 2 types of functions:
# 1 - Perform a task (above) 
# 2 - Calculate and return a value

# below is type 2 because it calculates and returns a value
round(1.9)


make the top function a type 2

def greet(name):
    print(f"Hi {name}")

def get_greeting(name):
    return f"Hi {name}"

message = get_greeting("Mosh")

# the reason to change it to type two is so it can be put into a variable and can be reused in the future

print(message)
file = open("content.txt", "w")
file.write(message)


def greet(name):
    print(f"Hi {name}")

print(greet("Mosh")) # --> returns Hi Mosh.... None --> None is used in an absence of a value.
# functions by default are None



