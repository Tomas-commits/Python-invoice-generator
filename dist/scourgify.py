import csv
import sys

if len(sys.argv) != 3:
    sys.exit("Only one argument allowed")
if not sys.argv[1].endswith(".csv") or not sys.argv[2].endswith(".csv"):
    sys.exit("not a csv file")
try:
    with open(sys.argv[1]) as file:
        reader = csv.reader(file)
        print(reader)


except FileNotFoundError:
    sys.exit("File not found")