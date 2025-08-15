# If you had a stack of books, the last book that was placed on the stack would be the first to be taken off.
# This bahavior is LIFO (Last In - First Out)
# Websites do this for when you press the back button. 
browsing_session = []
browsing_session.append(1)
browsing_session.append(2)
browsing_session.append(3)
print(browsing_sesssion)
last = browsing_session.pop()
print(last)
print(browsing_session)
print("redirect", browsing_session[-1])

if not browsing_session:
    print("disable")
  
# .append is used to add an item to the top of the stack
# .pop is used to remove the last item (top of the stack)
# use index of -1 returns the item on top of the stack

if not browsing_session:
    browsing_session[-1]:
















