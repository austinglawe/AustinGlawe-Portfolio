letters = ["a", "b", "c"]

for letter in letters:
    print(letter)

# tuple of two items - index, item at that index
for letter in enumerate(letters):
    print(letter)

for index, letter in enumerate(letters):
    print(index, letter)
