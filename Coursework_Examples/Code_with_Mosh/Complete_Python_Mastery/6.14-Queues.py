# Queues are FIFO (First In, First Out)
from collections imporet deque
queue = deque([])
queue.append(1)
queue.append(2)
queue.append(3)
queue.popleft()
print(queue)

if not queue:
    print("empty")
