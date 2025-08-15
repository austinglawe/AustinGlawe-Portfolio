def save_user(**user):
    print(user)

save_user(id=1, name="John", age=22)

# the double asterisk makes it into a dictionary

def save_user(**user):
    print(user["id"])
    print(user["name"])

save_user(id=1, name="John", age=22)
