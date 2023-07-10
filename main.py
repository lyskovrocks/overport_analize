import Service.regulars
from Service.OverportAnalizer import OverportAnalizer

overport = OverportAnalizer()

# overport.remove_duplicates()
# print(Service.regulars.RESIDENTIAL_COMPLEX)

overport.show_all_task_by_objects()


























# Считает кол-во заявок по ЖК.
# count = 0
# for task in range(1, 2000):
#     if "Южное" in overport.sheet[task][overport.ADDRESS].value:
#         count += 1
#         #print(overport.sheet[task][4].value)
# print("Заявок по ЖК София: ", count)
# stroka = 'Хэй, посмотри и попробуй найти тут ДуБль'
# stroka = stroka.lower()
# print(stroka.lower().find('дубль'))


# overport.show_tasks(100)
