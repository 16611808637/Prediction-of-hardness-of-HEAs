import itertools
import openpyxl

numbers = [0,8,10,12,14,15,16,18,20,22,24,25,26,28,30,32,34]
combinations = list(itertools.product(numbers, repeat=6))


valid_combinations = []
for combination in combinations:
    if sum(combination) == 100 and combination.count(0) >= 0 and combination.count(0) <= 2:
        valid_combinations.append(combination)


wb = openpyxl.Workbook()
ws = wb.active


for i, combination in enumerate(valid_combinations, start=1):
    ws.append(combination)

wb.save("1.xlsx")
print("have been generate 1.xlsx file")