# instead of writing out the function like this:
def sort_items(item):
    return item[1]

# you can use a lambda like this: 
items.sort(key=lambda item:item[1])
print(items)

# note: for a lambda you put 'parameters:expression'

# instead of defining a function once, we can use 'lambda' when we want to use an argument for another function.
