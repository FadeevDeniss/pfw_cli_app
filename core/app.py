import os

import argparse as ap
from typing import Union, Any
from operator import sub
from datetime import datetime

import pandas as pd
import openpyxl as op

from textwrap import dedent


class PersonalFinancialWallet:
    """
    Приложение командной строки "Персональный финансовый кошелек"

    Позволяет вести учет доходов и расходов и сохраняет информацию
    в файле xlsx
    ...

    Аттрибуты
    __________
    _prog_name : str
        имя приложения для отображения в консоли
    _prog_description : str
        описание приложения для отображения в консоли
    _actions : list
        список команд для приложения
    _root_dir : str
        строка с корневой директорией проекта
    _filename : str
        имя файла для хранения и записи информации
    columns : list
        список наименований для колонок
    categories : list
        список категорий для записей
    db_path : str
        строка содержащая полный путь до файла базы данных
    parser : ArgumentParser
        парсер аргументов командной строки
    """
    _prog_name = 'PFW'
    _prog_description = '''
        Личный Финансовый Кошелек
        ________________________________________________________________
        Приложение командной строки для ведения учета доходов и расходов
    '''

    _actions = ['balance', 'add', 'search', 'modify']

    _root_dir = os.getcwd()

    _filename = 'db.xlsx'

    columns = ['Category', 'Amount', 'Date', 'Description', 'Created at', 'Updated at']
    categories = ['Расход', 'Доход']

    def __init__(self):
        self.db_path = os.path.join(self._root_dir, f'database\\{self._filename}')
        self._create_db_if_not_exists()

        self.parser = ap.ArgumentParser(prog=self._prog_name,
                                        description=dedent(self._prog_description),
                                        formatter_class=ap.RawDescriptionHelpFormatter)

        self.parser.add_argument('action',
                                 choices=self._actions,
                                 help='Выполнить операцию одну на выбор - показать баланс,'
                                      ' добавить редактировать или удалить запись')
        self.parser.add_argument(
            '-c', '--cat', nargs='?', choices=self.categories, dest='category', help='Категория операции'
        )
        self.parser.add_argument('-a', '--amount', dest='amount', help='Сумма')
        self.parser.add_argument('--date', nargs='?', dest='date', help='Дата операции', const=datetime.date)
        self.parser.add_argument('-d', '--desc', nargs='?', dest='description', help='Дата операции')
        self.parser.add_argument('-i', '--index', nargs='?', type=int, dest='idx', help='Индекс записи')

    def display_balance(self, dataframe: pd.DataFrame) -> None:
        """
        Выводит в консоль информацию о балансе, расходы и доходы.

        :param dataframe: датафрейм с информацией о расходах и доходах
        :return: None
        """
        balance, income, outcome = self.count_balance(dataframe)
        if outcome > income:
            print(dedent(f'\n\nБюджет перерасходован!'))
        print(dedent(f'''
            __________________________
           | Текущий баланс: {balance:.2f} 
           | ------------------------- 
           | Доходы:         {income:.2f}          
           | ------------------------- 
           | Расходы:       {outcome:.2f}        
           | _________________________ 
        '''))

    def count_balance(self, dataframe: pd.DataFrame) -> tuple[int, int, int]:
        """
        Вычисляет баланс, расходы и доходы.

        :param dataframe: датафрейм с информацией о расходах и доходах
        :return: Кортеж с суммой баланса, общей суммой расходов и доходов
        """
        income_rows = self._filter_rows(dataframe, {'Category': 'Доход'})
        outcome_rows = self._filter_rows(dataframe, {'Category': 'Расход'})
        income = income_rows.Amount.sum() if not income_rows.empty else 0
        outcome = outcome_rows.Amount.sum() if not outcome_rows.empty else 0
        balance = sub(income, outcome)
        return balance, income, outcome

    def add_record(self, col_values: dict[str, Any]) -> None:
        """
        Добавляет новую запись в таблицу excel.

        :param col_values: Словарь со значениями ячеек,
                           переданных пользователем в аргументах командной строки
        :return: None
        """
        upd_datetime = datetime.now()
        col_values.update({'Created at': upd_datetime})
        self._save_to_excel(col_values)

        print(f'\nДобавлена новая запись:\n\n{" | ".join(str(i) for i in col_values.values())}')

    def modify_record(self, update_values: dict[str, Any], index: int) -> None:
        """
        Обновляет запись в файле excel.

        :param update_values: Словарь с новыми значениями ячеек
        :param index: Индекс записи для обновления
        :return: None
        """
        update_values['Updated at'] = datetime.now()
        self._save_to_excel(update_values, index)
        print(f'\nЗапись обновлена:\n{" | ".join(str(i) for i in update_values.values())}')

    def search_record(self, dataframe: pd.DataFrame, conditions: dict[str, str]) -> None:
        """
        Найти запись по значениям ячеек, переданных пользователем в аргументах
        командной строки.

        :param dataframe: датафрейм с информацией о расходах и доходах
        :param conditions: Словарь со значениями ячеек для поиска
        :return: None
        """
        self._filter_rows(dataframe, conditions, inplace=True)
        height, _ = dataframe.shape
        print(f'\nНайдено {height} записей\n_________________\n')
        if not dataframe.empty:
            print(dataframe)

    def start(self) -> None:
        """
        Запускает парсинг аргументов командной строки и вызывает
        методы класса в зависимости от команды, переданной пользователем.

        :return: None
        """
        args = self.parser.parse_args()
        _col_values = {}
        for attr in ('amount', 'category', 'date', 'description'):
            if getattr(args, attr):
                _col_values[attr.capitalize()] = getattr(args, attr)
        if args.action == 'balance':
            dataframe = self._load_dataframe_from_excel()
            if dataframe.empty:
                print('\nОшибка: Недостаточно данных для вычисления баланса.'
                      ' Отсутствуют доходы и расходы.')
                raise SystemExit(2)
            self.display_balance(dataframe)
        elif args.action == 'add':
            self.add_record(_col_values)
        elif args.action == 'search':
            self.search_record(self._load_dataframe_from_excel(), _col_values)
        else:
            self.modify_record(_col_values, args.idx)

    def _create_db_if_not_exists(self) -> None:
        """
        Создает файл excel в директории ./database/ если он не создан.

        :return: None
        """
        db_dirname = os.path.dirname(self.db_path)
        if not os.path.isdir(db_dirname):
            os.mkdir(db_dirname)
        if not os.path.isfile(self.db_path):
            wb = op.Workbook()
            sheet = wb.active
            sheet.title = 'Main'
            sheet.append(self.columns)
            wb.save(self.db_path)

    def _load_dataframe_from_excel(self) -> pd.DataFrame:
        """
        Загружает датафрейм из файла excel в директории ./database/

        :param path:
        :return: Датафрейм с данными о расходах и доходах из файла excel
        """
        return pd.read_excel(self.db_path, date_format='%Y-%m-%d', parse_dates=['Date'])

    def _save_to_excel(self, values: dict[str, Any], index: int = None) -> None:
        """
        Сохраняет переданные значения в файл excel.

        :param values: Словарь со значениями ячеек
        :param index: Индекс записи (По умолчанию равен None)
        :return: None
        """
        col_indexes = {i[1]: i[0] for i in enumerate(self.columns, 1)}
        workbook = op.load_workbook(self.db_path)
        sheet = workbook['Main']
        index = (sheet.max_row + 1) if index is None else (index + 2)
        for col_name, value in values.items():
            sheet.cell(row=index, column=col_indexes[col_name]).value = value
        workbook.save(self.db_path)

    @staticmethod
    def _filter_rows(dataframe: pd.DataFrame,
                     conditions: dict[str, str],
                     inplace: bool = False) -> Union[pd.DataFrame, None]:
        """
        Фильтрует датафрейм на основе переданных значений

        :param dataframe: Датафрейм с данными о расходах и доходах из файла excel
        :param conditions: Значения ячеек для фильтрации
        :param inplace: Опциональный аргумент, определяет создавать ли новый объект датафрейм,
                        или менять существующий. По умолчанию равен False
        :return: Отфильтрованный датафрейм или None, если записи не найдены
        """
        expr = ' & '.join([f'{item[0]} == "{item[1]}"' for item in conditions.items()])
        result = dataframe.query(expr, inplace=inplace)
        if not inplace:
            return result


