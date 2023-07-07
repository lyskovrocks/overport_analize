from OverportAnalizer import OverportAnalizer

# print(sheet.max_row)
# print(sheet[1][i].value)

overport = OverportAnalizer()
# overport.show_tasks(50)
print("Processing...")

# print(overport.sheet[32][4].value)

# Считает кол-во заявок по ЖК.
# count = 0
# for task in range(1, 2000):
#     if "Южное" in overport.sheet[task][overport.ADDRESS].value:
#         count += 1
#         #print(overport.sheet[task][4].value)
# print("Заявок по ЖК София: ", count)


overport.show_tasks(date_range_min='2023-05-01', date_range_max='2023-06-01')

