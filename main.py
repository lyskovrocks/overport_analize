from colorama import Fore
from OverportAnalizer import OverportAnalizer
import colorama


overport = OverportAnalizer()

for row in range(2,8):
    adress = overport.sheet[row][4].value
    task = overport.sheet[row][5].value
    status = overport.sheet[row][9].value
    comment = overport.sheet[row][6].value
    print(Fore.BLUE + status + Fore.RESET, adress, task,'\n' '         |||ОТВЕТ: ', Fore.YELLOW + str(comment) + Fore.RESET)


    #print(sheet[row][i].value)

    # print(sheet.max_row)
    # print(sheet[1][i].value)