# parameters can be added to functions by putting them into the parentheses
# parameters represent arguments
def greet(first_name, last_name):
    print(f"Hi {first_name} {last_name}")
    print("Welcome aboard")

greet("Mosh", "Hamedani")
# arguments are the actual values used within the parentheses
greet("John", "Smith")

# all arguments are REQUIRED
# you cannot do this:
greet("Mosh")
# it is missing 1 argument
