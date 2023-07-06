from openpyxl import __init__
import openpyxl
from colorama import Fore
from openpyxl.cell import read_only

book_origin = openpyxl.open("input_data/ads_journal.xlsx", read_only = True)
sheet = book_origin.active
i = 9
# print(sheet.max_row)
# print(sheet[1][i].value)

for row in range(2,8):
    adress = sheet[row][4].value
    task = sheet[row][5].value
    status = sheet[row][9].value
    comment = sheet[row][6].value
    print(Fore.BLUE + status + Fore.RESET, adress, task,'\n' '         |||ОТВЕТ: ', Fore.YELLOW + str(comment) + Fore.RESET)


    #print(sheet[row][i].value)