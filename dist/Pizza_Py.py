import sys
import csv
from tabulate import tabulate

if len(sys.argv) != 2:
    sys.exit("Only one argument allowed")
if not sys.argv[1].endswith(".csv"):
    sys.exit("not a csv file")

try:
    with open(sys.argv[1]) as file:
        for line, count in file, enumerate:
            print(line)
except FileNotFoundError:
    sys.exit()


# print(tabulate(table, headers, tablefmt="grid"))