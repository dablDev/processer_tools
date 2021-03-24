import io
import pandas as pd
import datetime
import xlsxwriter as X

from . import regions as r
XLSX_DATE_FORMAT = 'dd/mm/yy hh:mm:ss'


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


class CreationHelper:

    def __init__(self):
        self.output = io.BytesIO()
        self.workbook = X.Workbook(self.output)
        self.current_worksheet = None

    def SetSheet(self, name):
        assert isinstance(name, str)
        assert len(name) > 0
        assert len(name) < 30
        self.current_worksheet = self.workbook.add_worksheet(name=name)

    def WriteInt(self, cell, value, format):
        assert not (self.current_worksheet is None)
        assert isinstance(cell, CellPosition)
        assert isinstance(value, int)
        assert isinstance(format, X.workbook.Format) or (format is None)
        self.current_worksheet.write(cell.row, cell.column, value, format)

    def WriteFloat(self, cell, value, format):
        assert not (self.current_worksheet is None)
        assert isinstance(cell, CellPosition)
        assert isinstance(value, float) and not pd.isna(value)
        assert isinstance(format, X.workbook.Format) or (format is None)
        self.current_worksheet.write(cell.row, cell.column, value, format)

    def WriteComment(self, cell, comment, x_scale, y_scale):
        assert not (self.current_worksheet is None)
        assert isinstance(cell, CellPosition)
        assert isinstance(comment, str)
        assert isinstance(x_scale, float)
        assert isinstance(y_scale, float)
        self.current_worksheet.write_comment(cell.row, cell.column, comment, {'x_scale': x_scale, 'y_scale': y_scale})

    def WriteDate(self, cell, value, format):
        assert not (self.current_worksheet is None)
        assert isinstance(cell, CellPosition)
        assert isinstance(value, datetime.datetime)
        assert isinstance(format, X.workbook.Format) or (format is None)
        if isinstance(format, X.workbook.Format):
            format.set_num_format(XLSX_DATE_FORMAT)
        else:
            format = self.workbook.add_format({'num_format': XLSX_DATE_FORMAT})
        self.current_worksheet.write_datetime(cell.row, cell.column, value, format)

    def WriteStr(self, cell, value, format):
        assert not (self.current_worksheet is None)
        assert isinstance(cell, CellPosition)
        assert isinstance(value, str)
        assert isinstance(format, X.workbook.Format) or (format is None)
        self.current_worksheet.write(cell.row, cell.column, value, format)

    def WriteRegionDate(self):
        assert not (self.current_worksheet is None)
        format = self.workbook.add_format()
        format.set_bg_color('yellow')
        format.set_bold()
        self.WriteStr(cell=CellPosition(0,0), value='Регион', format=format)
        self.WriteStr(cell=CellPosition(0,1), value='Дата', format=format)

    def WriteRegionName(self, region_code):
        assert region_code in r.REGION_DECODER.keys()
        format = self.GetHeadFormat()
        self.WriteStr(cell=CellPosition(0,0), value=r.REGION_DECODER[region_code], format=format)

    def GetHeadFormat(self):
        res = self.GetFormatObj()
        res.set_bg_color('yellow')
        res.set_bold()
        return res

    def GetFormatObj(self):
        return self.workbook.add_format()

    def SetOutputName(self, name):
        assert isinstance(name, str)
        assert len(name) > 0
        self.output.name = name

    def WriteTitles(self, title_list, title_row):
        assert not(self.current_worksheet is None)
        assert isinstance(title_list, list)
        assert len(title_list) > 0
        assert isinstance(title_row, int)
        assert title_row >= 0
        format = self.GetFormatObj()
        format.set_bold()
        for i, title in enumerate(title_list):
            assert isinstance(title, str)
            self.WriteStr(cell=CellPosition(title_row, i), value=title, format=format)

    def GetOutput(self):
        self.workbook.close()
        return io.BytesIO(self.output.getvalue())

    def AdjustColumns(self, start_col, end_col, width):
        assert not (self.current_worksheet is None)
        assert isinstance(start_col, int)
        assert isinstance(end_col, int)
        assert isinstance(width, int)
        assert start_col >= 0
        assert end_col > 0
        assert start_col <= end_col
        assert width > 0

        self.current_worksheet.set_column(start_col, end_col, width)


    def MergeRange(self, first_row, first_col, last_row, last_col, data, cell_format):
        assert isinstance(first_row, int)
        assert isinstance(first_col, int)
        assert isinstance(last_row, int)
        assert isinstance(last_col, int)
        assert first_row >= 0
        assert first_col >= 0
        assert last_row >= 0
        assert last_col >= 0
        assert not (data is None)
        assert (cell_format is None) or isinstance(cell_format, X.workbook.Format)
        self.current_worksheet.merge_range(first_row, first_col, last_row, last_col, data, cell_format)

