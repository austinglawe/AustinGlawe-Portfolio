# While Loops are used while conditions are true
number = 100
while number > 0:
    print(number)
    number //= 2

command = ""
while command != "quit":
    command = input(">")
    print("ECHO", command")


# an amateur may do (DO NOT DO THIS):
command = ""
while command != "quit" or command != "quit":
    command = input(">")
    print("ECHO", command")

# The better way to do it:
command = ""
while command.lower() != "quit":
    command = input(">")
    print("ECHO", command")
          









