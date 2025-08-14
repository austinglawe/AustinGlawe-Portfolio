# you could concatenate like this:
first = "Mosh"
last = "Hamedani"
full = first + " " + last
print(full)

# But instead you should use formatted strings like this:
first = "Mosh"
last = "Hamedani"
full = f"{first} {last}"
print(full)

# you can put any value expressions in the curly braces
first = "Mosh"
last = "Hamedani"
full = f"{len(first)} {2 + 2}"
print(full)
