import logging

import openpyxl
from colorama import Fore
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

import Service.regulars as regulars


class OverportAnalizer:
    DATE = 3
    ADDRESS = 4
    TASK = 5
    STATUS = 9
    COMMENT = 6
    START_ROW = 2
    ORIGINAL_SHEET_NAME = 'Лист 1'
    NEW_SHEET_NAME = 'Without_Duplicates'


    def __init__(self, path = None):

        if path is None:
            raise AttributeError('Укажите имя файла в конструкторе')

        self.wb: Workbook = openpyxl.load_workbook(path, read_only=False)

        # self.wb.remove(self.NEW_SHEET_NAME)
        self.wb.create_sheet(self.NEW_SHEET_NAME, 0)

        self.wb.save(path)
        self.path = path
        self.sheet: Worksheet = self.wb.active
        self.sheet_origin: Worksheet = self.wb[self.ORIGINAL_SHEET_NAME]


    def remove_duplicates(self):
        logging.info('Removing duplicates...')
        deleted_row_count = 0
        for task in self.sheet_origin.iter_rows(
                min_row=self.START_ROW,
                max_row=self.sheet_origin.max_row,
                values_only=True):
            if task[self.COMMENT] is None:
                self.sheet.append(task)

            else:
                comment = str(task[self.COMMENT]).lower()

                duplicates_flag = False
                for word in regulars.DUPLICATES:
                    if comment.find(word) >= 0:
                        duplicates_flag = True
                        break

                if duplicates_flag:
                    deleted_row_count += 1
                    continue

                self.sheet.append(task)

        self.wb.save(self.path)
        print(f'Удалено {deleted_row_count} дублей')


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
    # Прочитать книгу
    # отсортировать данные по ЖК и кол-ву дублей
    # записать данные в новую книгу без дублей
    # Сослаться на новую книгу во всех последующих функциях