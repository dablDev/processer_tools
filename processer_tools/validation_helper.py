import copy
import io
import pandas as pd
import xlrd
import datetime
import re
from . import regions as r

KS_KSG_TYPE = 'ks'
DS_KSG_TYPE = 'ds'

REGIONS = r.REGIONS_DICT['Encoder']

class CellPosition:

    def __init__(self, row_number, column_number):
        assert isinstance(column_number, int)
        assert column_number >= 0
        assert isinstance(row_number, int)
        assert row_number >= 0
        self.column = column_number
        self.row = row_number

    def GetString(self):
        assert self.column <= 26
        column = chr( (self.column + 1) + 64)
        return column + str(self.row + 1)


def GetValue(df, cell):
    assert isinstance(df, pd.DataFrame)
    assert isinstance(cell, CellPosition)
    res = df.values[cell.row][cell.column]
    if isinstance(res, str):
        res = res.strip()
    return res

def IsDate(value):
    return isinstance(value, datetime.datetime)

def IsType(type_, value):
    assert not (type_ is None)
    assert not (value is None)
    assert isinstance(type_, type)
    return isinstance(value, type_)

def IsSameString(actual_string, must_be_string):
    return actual_string == must_be_string

def GetINN(df, cell):
    assert isinstance(df, pd.DataFrame)
    assert isinstance(cell, CellPosition)
    value = GetValue(df, cell)
    assert isinstance(value, (int, str))
    if type(value) == int:
        assert len(str(value)) == 10 or len(str(value)) == 12
        res = str(value)
    else:
        value = value.strip()
        assert len(str(value)) == 10 or len(str(value)) == 12
        res = value
    return res

def RegExMatch(target_str, regex_str):
    pattern = re.compile(regex_str)
    if pattern.fullmatch(target_str) is None:
        result = False
    else:
        result = True
    return result


def VerifyKSGCode(code, kindOfKSG):
    assert (kindOfKSG == KS_KSG_TYPE) or (kindOfKSG == DS_KSG_TYPE)
    done = False
    error = ''
    if type(code) != str:
        error = 'Неверный тип данных кода. Должен быть строкой, сейчас {}'.format(type(code))
    if kindOfKSG == KS_KSG_TYPE:
        code_prefix = 'st'
    else:
        code_prefix = 'ds'
    if (error == '') and not (len(code) == 4 or len(code) == 8 or len(code) == 12 or len(code) == 14):
        error = 'Длина текущего кода {}, сам код {}. Длина кода может быть равна 4 символам, например {}01, 8 символам, например {}29.004, либо 12 символам (с региональным расширением), например {}29.004.001 ' \
                'либо 14 символам, например {}29.004.001.1' \
                .format(len(code), code, code_prefix, code_prefix, code_prefix, code_prefix)
    if error == '':
        if not (code[:2] == code_prefix):
            error = 'Код должен начинаться с {}. Начинается с {}'.format(code_prefix, code[:2])
    if error == '':
        if not (code[2] in ['0', '1', '2', '3', '4']):
            error = 'Невозможный код. Третий символ должен быть от 0 до 4'
    if error == '':
        if not (code[3] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']):
            error = 'Невозможный код. Четвертый символ должен быть от 0 до 9'
    if error == '' and ( len(code) > 4):
        if not (code[4] == '.'):
            error = 'Невозможный код. Пятый символ кода должен быть точкой .'
        if error == '':
            if not (code[5] == '0'):
                error = 'Невозможный код. Шестой символ кода должен быть 0'
        if error == '':
            if not (code[6] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] ):
                error = 'Невозможный код. Седьмой символ кода должен быть от 0 до 9'
        if error == '':
            if not (code[7] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] ):
                error = 'Невозможный код. Восьмой символ кода должен быть от 0 до 9'
    if (error == '' ) and (len(code) == 12):
        if error == '':
            if not (code[8] == '.'):
                error = 'Невозможный код. Девятый символ кода должен быть точкой .'
        if error == '':
            if not (code[9] == '0'):
                error = 'Невозможный код. Десятый символ кода должен быть 0'
        if error == '':
            if not (code[10] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']):
                error = 'Невозможный код. Одиннадцатый символ кода должен быть от 0 до 9'
        if error == '':
            if not (code[11] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']):
                error = 'Невозможный код. Двенадцатый символ кода должен быть от 0 до 9'
    if (error == '') and (len(code) == 14):
        if error == '':
            if not (code[8] == '.'):
                error = 'Невозможный код. Девятый символ кода должен быть точкой .'
        if error == '':
            if not (code[9] == '0'):
                error = 'Невозможный код. Десятый символ кода должен быть 0'
        if error == '':
            if not (code[10] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']):
                error = 'Невозможный код. Одиннадцатый символ кода должен быть от 0 до 9'
        if error == '':
            if not (code[11] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']):
                error = 'Невозможный код. Двенадцатый символ кода должен быть от 0 до 9'
        if error == '':
            if not (code[12] == '.'):
                error = 'Невозможный код. Тринадцатый символ кода должен быть точкой .'
        if error == '':
            if not (code[13] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']):
                error = 'Невозможный код. Четырнадцатый символ кода должен быть от 0 до 9'
    if error == '':
        done = True
    return done, error



class ValidationHelper:

    def __init__(self, template_io):
        assert isinstance(template_io, io.BytesIO)
        try:
            excel_file = pd.ExcelFile(io=copy.deepcopy(template_io))
            self.sheet_names =excel_file.sheet_names
            self.error = ''
            self.sheets = []
            for i in range(len(self.sheet_names)):
                self.sheets.append(excel_file.parse(i, header=None))
        except xlrd.biffh.XLRDError:
            self.sheet_names = None
            self.sheets = []
            self.error = 'Приложен не эксель файл или битый эксель файл'
        self.current_sheet = None
        self.current_sheet_name = None

    def _get_sheet(self, sheet_name):
        assert sheet_name in self.sheet_names
        assert self.NoError()
        i = 0; found = False; found_idx = 0
        while (i < len(self.sheet_names)) and not found:
            found = self.sheet_names[i] == sheet_name
            found_idx = i
            i = i + 1
        assert found
        return self.sheets[found_idx], self.sheet_names[found_idx]

    def GetValue(self, cell):
        assert isinstance(cell, CellPosition)
        if self.NoError():
            res = GetValue(self.current_sheet, cell)
            if isinstance(res, str):
                res = res.strip()
        else:
            res = None
        return res


    def SetSheet(self, sheet_name):
        if self.NoError():
            self.current_sheet, self.current_sheet_name = self._get_sheet(sheet_name)

    def HasSheet(self, sheet_name):
        assert isinstance(sheet_name, str)
        if self.NoError():
            res = sheet_name in self.sheet_names
        else:
            res = False
        return res

    def GetSheetNames(self):
        if self.NoError():
            res = self.sheet_names
        else:
            res = []
        return res

    def CheckTitle(self, title_list, row_number):
        assert isinstance(row_number, int)
        assert row_number >= 0
        assert isinstance(title_list, list)
        assert len(title_list) > 0
        for elem in title_list:
            assert isinstance(elem, str) or pd.isna(elem)
        if self.NoError():
            for i, title in enumerate(title_list):
                if pd.isna(title):
                    self.IsEmpty(cell=CellPosition(row_number, i))
                else:
                    self.IsSameString(cell=CellPosition(row_number, i), must_be_string=title)
        self.TableColumnsCheck(len(title_list))

    def CheckTitleRegEx(self, regex_list, row_number):
        assert isinstance(row_number, int)
        assert row_number >= 0
        assert isinstance(regex_list, list)
        assert len(regex_list) > 0
        for regex_elem in regex_list:
            assert isinstance(regex_elem, str) or pd.isna(regex_elem)
        if self.NoError():
            for i, regex_pattern in enumerate(regex_list):
                if pd.isna(regex_pattern):
                    self.IsEmpty(cell=CellPosition(row_number, i))
                else:
                    self.SatisfiesRegex(cell=CellPosition(row_number, i), regex_pattern=regex_pattern)

    def NoError(self):
        return len(self.error) == 0

    def SheetsLenCheck(self, sheets_len):
        assert isinstance(sheets_len, int)
        if self.NoError():
            if len(self.sheet_names) != sheets_len:
                self.error += 'Неверное число листов в файле эксель. Прикреплено {}, должно быть {}'.format(len(self.sheet_names), sheets_len)

    def HasSheets(self, sheets):
        assert isinstance(sheets, list)
        assert len(sheets) > 0
        if self.NoError():
            sheet_names = self.sheet_names
            for req_sheet in sheets:

                i = 0; found = False
                while (i < len(sheet_names)) and not found:
                    found = sheet_names[i] == req_sheet
                    i = i + 1

                if not found:
                    self.error += 'Не обнаружен лист {}'.format(req_sheet)

    def SheetsNamesCheck(self, sheet_names):
        assert isinstance(sheet_names, list)
        if self.NoError():
            excel_names = self.sheet_names
            i = 0; found_error = len(sheet_names) != len(excel_names)
            while (i < len(excel_names)) and not found_error:
                assert isinstance(sheet_names[i], str)
                found_error = sheet_names[i] != excel_names[i]
                i = i + 1
            if found_error:
                self.error += 'Не совпадают названия листов в экселе. Сейчас {}, должно быть {}'.format(excel_names, sheet_names)

    def BlancTableCheck(self):
        if self.NoError():
            actual_cols = len(self.current_sheet.columns)
            if not (actual_cols == 0):
                self.error += 'Лист {} должен быть пустым, сейчас он содержит колонки {}\n'.format(
                    self.current_sheet_name,
                    self.current_sheet.columns)

    def TableColumnsCheck(self, columns_num):
        assert isinstance(columns_num, int); assert columns_num > 0
        if self.NoError():
            actual_cols = len(self.current_sheet.columns)
            if not (actual_cols == columns_num):
                self.error += 'Количество колонок в листе {} должно быть: {}, сейчас {}\n'.format(self.current_sheet_name,
                                                                                                  columns_num,
                                                                                                  actual_cols)

    def TableRowsCheck(self, rows_num):
        assert isinstance(rows_num, int); assert rows_num > 0
        if self.NoError():
            actual_rows = self.current_sheet.shape[0]
            if not (actual_rows - 1 == rows_num):
                self.error += 'Количество строк в таблице {} должно быть: {}, сейчас {}\n'.format(self.current_sheet_name,
                                                                                                rows_num,
                                                                                                actual_rows - 1)

    def CheckUniqueOrEmpty(self, column_number, start_row):
        assert isinstance(column_number, int)
        assert column_number >= 0
        assert isinstance(start_row, int)
        assert start_row >= 0
        if self.NoError():
            rows_number = self.GetRows()
            values = []
            for i in range(start_row, rows_number):
                value = self.GetValue(CellPosition(i, column_number))
                if not pd.isna(value):
                    if value in values:
                        self.error += 'Значение {} повторяется повторно в ячейке {}, хотя все значения в этой колонке ' \
                                      'должны быть уникальны'.format(value, CellPosition(i, column_number).GetString())
                    else:
                        values.append(value)



    def TableSizeCheck(self, columns_num, rows_num):
        assert isinstance(columns_num, int); assert columns_num > 0
        assert isinstance(rows_num, int); assert rows_num > 0
        if self.NoError():
            assert not (self.current_sheet is None)
            assert not (self.current_sheet_name is None), self.current_sheet_name
            self.TableColumnsCheck(columns_num=columns_num)
            self.TableRowsCheck(rows_num=rows_num)

    def GetRows(self):
        if self.NoError():
            actual_rows = self.current_sheet.shape[0]
        else:
            actual_rows = 0
        return actual_rows

    def RegionDateCheck(self):
        if self.NoError():
            region = GetValue(self.current_sheet, CellPosition(0, 0))
            date = GetValue(self.current_sheet, CellPosition(0, 1))
            if not (region in REGIONS.keys()):
                self.error += "Недопустимый регион. В ячейке {} {}, должно быть что-то из списка: {}\n".format(CellPosition(0, 0).GetString(),
                                                                                                               region,
                                                                                                               list(REGIONS.keys()) )
            if not IsDate(date):
                self.error += 'Дата должна быть записана в ячейке {} формате ДД.ММ.ГГГГ, пример: 31.12.2019 (новый год). Сейчас {}\n'.format(CellPosition(0, 1).GetString(),
                                                                                                                                             date)

    def IsID(self, cell):
        assert isinstance(cell, CellPosition)
        if self.NoError():
            value = GetValue(self.current_sheet, cell)
            if not pd.isna(value) and not isinstance(value, int):
                self.error += 'Значение {} в ячейке {}. Она должна быть либо пустой, либо должна содержать целое значение'.format(value, cell.GetString())


    def IsEmpty(self, cell):
        assert isinstance(cell, CellPosition)
        if self.NoError():
            value = GetValue(self.current_sheet, cell)
            if not pd.isna(value):
                self.error += 'Ячейка {} должна быть пуста. Сейчас там {}'.format(cell.GetString(), value)


    def IsDate(self, cell):
        assert isinstance(cell, CellPosition)
        if self.NoError():
            date = GetValue(self.current_sheet, cell)
            if not IsDate(date):
                self.error += 'Дата должна быть записана в ячейке {} формате ДД.ММ.ГГГГ, пример: 31.12.2019 (новый год). Сейчас {}\n'.format(cell.GetString(),
                                                                                                                                             date)

    def IsSameString(self, cell, must_be_string):
        assert isinstance(cell, CellPosition)
        assert isinstance(must_be_string, str)
        if self.NoError():
            actual_string = GetValue(self.current_sheet, cell)
            if not IsSameString(actual_string, must_be_string):
                self.error += 'Ячейка {} должна иметь значение: {}. Вместо этого значение: {} '.format(
                    cell.GetString(),
                    must_be_string,
                    actual_string
                )

    def IsSameNumber(self, cell, must_be_value):
        assert isinstance(cell, CellPosition)
        assert isinstance(must_be_value, (int, float))
        if self.NoError():
            actual_value = GetValue(self.current_sheet, cell)
            if not (actual_value == must_be_value):
                self.error += 'Ячейка {} должна иметь значение: {}. Вместо этого значение: {} '.format(
                    cell.GetString(),
                    must_be_value,
                    actual_value
                )


    def IsFloat(self, cell):
        assert isinstance(cell, CellPosition)
        if self.NoError():
            value = GetValue(self.current_sheet, cell)
            if (not isinstance(value, (float, int)) )or pd.isna(value):
                self.error = 'Ячейка {} должна быть числом. Сейчас {}'.format(cell.GetString(), type(value))

    def IsInt(self, cell):
        assert isinstance(cell, CellPosition)
        if self.NoError():
            value = GetValue(self.current_sheet, cell)
            if not isinstance(value, int):
                self.error = 'Ячейка {} должна быть  целым числом. Сейчас {}'.format(cell.GetString(), type(value))

    def IsStr(self, cell):
        assert isinstance(cell, CellPosition)
        if self.NoError():
            value = GetValue(self.current_sheet, cell)
            if not isinstance(value, str):
                self.error = 'Ячейка {} должна быть строкой. Сейчас {}'.format(cell.GetString(), type(value))

    def IsINN(self, cell):
        assert isinstance(cell, CellPosition)
        if self.NoError():
            value = GetValue(self.current_sheet, cell)

            if type(value) == int:
                if len(str(value)) != 10 and len(str(value)) != 12:
                    self.error = 'Недопустимая длина ИНН в ячейке {}, равная {} (должна быть 10 или 12). Введенное' \
                                 'значение {}'.format(cell.GetString(), len(str(value)), str(value))
            elif type(value) == str:
                value = value.strip()
                if len(value) != 10 and len(value) != 12:
                    self.error = 'Недопустимая длина ИНН в ячейке {}, равная {} (должна быть 10 или 12). Введенное' \
                                 'значение {}'.format(cell.GetString(), len(str(value)), str(value))
            else:
                self.error = 'Недопустимый тип данных {}. Введенное значение {}, ячейка {}'.format(type(value), value, cell.GetString())

    def InList(self, cell, values_list):
        assert isinstance(cell, CellPosition)
        assert isinstance(values_list, list)
        assert len(values_list) > 0
        if self.NoError():
            value = GetValue(self.current_sheet, cell)
            if not value in values_list:
                self.error = 'Ячейка {} должна иметь одно из следующих значений {}. Сейчас {}'.format(cell.GetString(),
                                                                                                      values_list,
                                                                                                      value)


    def NotInList(self, cell, values_list):
        assert isinstance(cell, CellPosition)
        assert isinstance(values_list, list)
        if self.NoError():
            value = GetValue(self.current_sheet, cell)
            if value in values_list:
                self.error = 'Ячейка {} не должна быть среди следующих значений {}. Сейчас {}'.format(cell.GetString(),
                                                                                                      values_list,
                                                                                                      value)

    def SatisfiesRegex(self, cell, regex_pattern):
        assert isinstance(cell, CellPosition)
        assert isinstance(regex_pattern, str)
        if self.NoError():
            value = GetValue(self.current_sheet, cell)
            if not RegExMatch(str(value), regex_pattern):
                self.error = 'Ячейка {}, содержащая значение {}, должна удовлетворять регулярному выражению {}'.format(
                    cell.GetString(),
                    value,
                    regex_pattern
                )

