letters = ["a", "b", "c"]
print(letters.index("a"))
print(leeters.index("d")) # ValueError: 'd' is not in list.

# first check if it is in the list
if "d" in letters:
    print(letters.index("d"))

print(letters.count("d"))
