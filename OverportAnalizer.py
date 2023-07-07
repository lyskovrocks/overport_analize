import shutil
import openpyxl
from colorama import Fore
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class OverportAnalizer:
    DATE = 3
    ADDRESS = 4
    TASK = 5
    STATUS = 9
    COMMENT = 6

    def __init__(self):
        backup_path = self.backup_sheet("input_data/ads_journal.xlsx")
        self.wb: Workbook = openpyxl.load_workbook(backup_path)
        self.sheet: Worksheet = self.wb.active

    def backup_sheet(self, path):
        shutil.copyfile(path, path + '.backup.xlsx')
        return path + '.backup.xlsx'

    def show_tasks(self, task_quantity=all, date_range_min='2020-01-01', date_range_max='2100-01-01'):
        # Определяем нужное кол-во тасков
        if task_quantity == all:
            tq = self.sheet.max_row
        else:
            tq = task_quantity

        # Определяем временно диапазон
        for row in range(2, tq + 2):
            if date_range_min < self.sheet[row][self.DATE].value < date_range_max:
                print(f"{Fore.BLUE + self.sheet[row][self.STATUS].value + Fore.RESET} "
                      f"{self.sheet[row][self.DATE].value} "
                      f"{self.sheet[row][self.ADDRESS].value} "
                      f"{self.sheet[row][self.TASK].value}"
                      f"{self.sheet[row][self.COMMENT].value}")

    def show_redcom_tasks(self):
        ...

    # todo Удаление дублей с подсчётом их кол-ва на ЖК и возвращением их кол-ва и процента от общего кол-ва заявок.
    #  Создание нового файла excel без дублей, с которым будут работать все функции

    def remove_duplicates(self):
        ...
