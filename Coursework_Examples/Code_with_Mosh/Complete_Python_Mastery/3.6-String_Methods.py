course = "  pythoN   Programming  "

# Everything in python is an object and all functions of objects are called methods. Therefore, we use methods by using the dot notation to gain access to them.
print(course.upper())
print(course.lower())
print(course.title())
print(course.strip())
print(course.lstrip())
print(course.rstrip())
# for .find - pass an argument through
print(course.find("pro"))
print(course.replace("p", "j"))
# expression is a piece of code that produces a value. Below is an expression - a Boolean (True or False)
print("pro" in course)
# not operator:
print("swift" not in course)
