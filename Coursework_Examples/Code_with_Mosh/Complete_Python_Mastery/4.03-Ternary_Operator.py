age = 22
if age >= 18:
    print("Eligible")
else:
    print("Not Eligible")

# Alternatively you can do this, to make it cleaner:
age = 22
if age >= 18:
    message = "Eligible"
else:
    message = "Not Eligible"

print(message)

# To make it even cleaner you can do this - this is a ternary operator:
age 22
message = "Eligible" if age >= 18 else "Not Eligible"
print(message)
