# scope - region of the code where a variable is defined.
def greet(name):
    message = "a"

print(message)
# NameError: name 'message' is not defined
# the scope of the 'name' and 'message' variables are the greet function.
# you refer to them as local variables. They are only local to the function, but do not exist outside of them.

# This allows multiple functions to have the same variable names like below:
def greet(name):
    message = "a"

def send_email(name):
    message = "b"

greet("Mosh")

# we also have global variables - message becomes a global variable below
# it can be used anywhere within the file
message = "a"
def greet(name):
    
def send_email(name):
    message = "b"

send_email("Mosh")
print(message) # prints a

# best practice is to have local variables within a function











