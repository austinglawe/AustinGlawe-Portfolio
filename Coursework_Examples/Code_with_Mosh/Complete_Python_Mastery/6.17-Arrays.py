[1, 2, 3]
# Arrays are best for large amounts of numbers - 10,000+

from array import array

#type codes
numbers = array("i", [1, 2, 3])
numbers.append(4)
numbers.insert()
numbers.pop()
numbers.remove()
numbers[0] = 1.0 #TyperError: integer argument expected, got float
# ("i",    determines the type code (in this case integer)


