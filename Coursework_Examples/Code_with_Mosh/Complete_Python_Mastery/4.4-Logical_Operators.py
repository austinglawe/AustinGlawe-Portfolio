# and
# or
# not

# and
# Only if both conditions are True, then it will be True
high_income = True
good_credit = True

if high_income and good_credit:
    print("Eligible")
else:
    print("Not Eligible")

# or
# If either of the conditions is True, then it will be True
high_income = True
good_credit = False

if high_income or good_credit:
    print("Eligible")
else:
    print("Not Eligible")

# not
# inverses the value of a variable
high_income = True
good_credit = False
student = True

if not student:
    print("Eligible")
else:
    print("Not Eligible")


# use of all 3
high_income = True
good_credit = False
student = True

if (high_income or good_credit) and not student:
    print("Eligible")
else:
    print("Not Eligible")









