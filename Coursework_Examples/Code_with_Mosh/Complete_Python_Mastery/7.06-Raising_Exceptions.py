def calculate_xfactor(age):
    if age <= 0:
        raise ValueError("Age cannot be 0 or less.")
    return 10 / age

# lookup python 3 built in exceptions

calculate_xfactor(-1)

try:
    calculate_xfactor(-1)
except ValueError as error:
    print(error)







