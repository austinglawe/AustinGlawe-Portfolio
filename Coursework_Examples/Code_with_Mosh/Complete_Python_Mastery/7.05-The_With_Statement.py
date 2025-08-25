try:
    with open("app.py") as file:
        print("File opened.")
        #file.__enter
        #file.__exit
    age = int(input("Age: "))
    xfactor = 10 / age
except (ValueError, ZeroDivisionEror):
    print("You didn't enter a valid age.")
else:
    print("No exceptions were thrown.")

# no longer needed:
#finally:
#    file.close()





try:
    with open("app.py") as file, open("another.txt) as target:
        print("File opened.")
    age = int(input("Age: "))
    xfactor = 10 / age
except (ValueError, ZeroDivisionEror):
    print("You didn't enter a valid age.")
else:
    print("No exceptions were thrown.")









