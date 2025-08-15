def increment(number, by=1):
    return number + by

print(increment(2))
# 1 is used by default for the 'by' argument because it is not present

def increment(number, by=1):
    return number + by

print(increment(2, 5))

# ALL OPTIONAL PARAMETERS SHOULD COME AFTER THE REQUIRED ONES
# below would not work..
def increment(number, by=1, another):
    return number + by
