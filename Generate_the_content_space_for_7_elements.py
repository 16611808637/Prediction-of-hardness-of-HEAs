import itertools
import openpyxl

numbers = [0,5,8,10,12,15,18,20,22,25,28,30,32,35]
combinations = list(itertools.product(numbers, repeat=7))


valid_combinations = []
for combination in combinations:
    if sum(combination) == 100 and combination.count(0) >= 1 and combination.count(0) <= 3:
        valid_combinations.append(combination)


wb = openpyxl.Workbook()
ws = wb.active


for i, combination in enumerate(valid_combinations, start=1):
    ws.append(combination)

wb.save("1.xlsx")
print("have been generate 1.xlsx file")