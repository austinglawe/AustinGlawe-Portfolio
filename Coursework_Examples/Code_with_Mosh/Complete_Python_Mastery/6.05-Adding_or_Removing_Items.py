letters = ["a", "b", "c"]

# Add Items to a list
# to the end - .append
letters.append("d")
print(letters)

# to a specific spot - .insert
letters.insert(0, "-")
print(letters)

# Remove
# at the end - .pop
letters.pop()
print(letters)

# at a given index - .pop(_)
letters.pop(0)
print(letters)

# if you do not know the index
letters.remove("b") # removes the first occurrance of 'b'
print(letters)

del letters[0:2] # del can delete a range
print(letters)

# remove all items in the list - .clear
letters.clear
print(letters)






