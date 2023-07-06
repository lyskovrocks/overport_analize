import shutil
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class OverportAnalizer:
    ADDRESS = 4
    TASK = 5
    STATUS = 9
    COMMENT = 6


    def __init__(self):
        backup_path = self.backup_sheet("input_data/ads_journal.xlsx")
        self.wb:Workbook = openpyxl.load_workbook(backup_path)
        self.sheet:Worksheet = self.wb.active


    def backup_sheet(self, path):
        shutil.copyfile(path, path + '.backup.xlsx')
        return path + '.backup.xlsx'

    def show_rows(self, rows_count):
        for row in range(1, rows_count+1):
            print(f"{self.sheet[row][self.ADDRESS].value} "
                  f"{self.sheet[row][self.TASK].value} "
                  f"{self.sheet[row][self.STATUS].value}")



if __name__ == "__main__":
    overport = OverportAnalizer()
    overport.show_rows(6)
    print('end')