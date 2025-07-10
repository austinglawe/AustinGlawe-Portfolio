# -----------------------------------------
# Python_File_IO:
# Basic file input/output
# -----------------------------------------
#
# Open file for reading:
#   file = open("example.txt", "r")
#   content = file.read()
#   file.close()
#
# Write to file:
#   file = open("output.txt", "w")
#   file.write("Hello, world!")
#   file.close()
#
# Using 'with' statement (recommended):
#   with open("example.txt", "r") as file:
#       content = file.read()
#
# Read lines:
#   with open("example.txt", "r") as file:
#       for line in file:
#           print(line.strip())
#
# Write multiple lines:
#   lines = ["Line 1\n", "Line 2\n", "Line 3\n"]
#   with open("output.txt", "w") as file:
#       file.writelines(lines)
#
# Best practices:
# - Use 'with' to handle files safely.
# - Specify mode explicitly.
# - Strip lines when reading.
# - Handle file-related exceptions.
#
# -----------------------------------------
