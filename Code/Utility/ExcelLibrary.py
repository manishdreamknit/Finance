import os
import sys
import   xlrd
from datetime import date, datetime, timedelta
from xlrd import open_workbook, xldate_as_tuple, \
    XL_CELL_NUMBER, XL_CELL_DATE, XL_CELL_TEXT, cellname
from xlwt import easyxf, XFStyle, Workbook
from xlutils.copy import copy

"""
This test library provides some keywords to allow
opening, reading, writing, and saving Excel files
from Robot Framework
"""

class ExcelLibrary:

    VERSION = '0.0.1'

    def __init__(self, slash='/'):
        self.wb = None
        self.tb = None
        self.slash = slash
        self.tmpDir = os.path.abspath(os.path.join( __file__, os.path.pardir))

    def open_excel(self, fname):
        """Open the Excel file indicated by fname"""
        print 'Opening file at %s' % fname
        self.wb = open_workbook(os.path.join(self.slash, self.tmpDir, fname), formatting_info=True)#, on_demand=True)


    def read_cell(self, row, column, sheetname):
        """Return the value stored in the cell indicated by
        row and column.
        """
        sheet = self.wb.sheet_by_name(sheetname)
        cv = sheet.cell(int(row), int(column)).value
        print 'Cell %s!' % cv
        return cv

    def put_number_to_cell(self, row, column, value, sheetname):
        """Sets the value of the indicated cell to be
        the number given in the parameter.
        """
        if self.wb:
            cell = self.wb.sheet_by_name(sheetname).cell(int(row), int(column))
            if cell.ctype == XL_CELL_NUMBER:
                self.wb.sheets()
                if not self.tb:
                    self.tb = copy(self.wb)
        if self.tb:
            plain = easyxf('')
            self.tb.sheet_by_name(sheetname).write(int(row),
                                       int(column),
                                       float(value),
                                       plain)

    def put_string_to_cell(self, row, column, value):
        """Sets the value of the indicated cell to be
        the string given in the parameter.
        """
        if self.wb:
            cell = self.wb.get_sheet(0).cell(int(row), int(column))
            if cell.ctype == XL_CELL_TEXT:
                self.wb.sheets()
                if not self.tb:
                    self.tb = copy(self.wb)
        if self.tb:
            plain = easyxf('')
            self.tb.get_sheet(0).write(int(row),
                                       int(column),
                                       value,
                                       plain)

    def put_date_to_cell(self, row, column, value, dateFrm='d.M.yyyy'):
        """Sets the value of the indicated cell to be
        the date given in the parameter. The format of the resulting
        date may be given, too.
        """
        if self.wb:
            cell = self.wb.get_sheet(0).cell(int(row), int(column))
            if cell.ctype == XL_CELL_DATE:
                self.wb.sheets()
                if not self.tb:
                    self.tb = copy(self.wb)
        if self.tb:
            print(value)
            dt = value.split('.')
            dti = [int(dt[2]), int(dt[1]), int(dt[0])]
            print(dt, dti)
            ymd = datetime(*dti)
            plain = easyxf('', num_format_str=dateFrm)
            self.tb.get_sheet(0).write(int(row),
                                       int(column),
                                       ymd,
                                       plain)

    def modify_cell_with(self, row, column, op, val):
        """Modifies a number cell
        with the given operation and value.
        """
        cell = self.wb.get_sheet(0).cell(int(row), int(column))
        curval = cell.value
        if cell.ctype == XL_CELL_NUMBER:
            self.wb.sheets()
            if not self.tb:
                self.tb = copy(self.wb)
            plain = easyxf('')
            modexpr = str(curval)+op+val
            self.tb.get_sheet(0).write(int(row),
                                       int(column),
                                       eval(modexpr),
                                       plain)

    def add_to_date(self, row, column, numdays):
        """Adds a number of days to the
        date in the indicated cell.
        """
        cell = self.wb.get_sheet(0).cell(int(row), int(column))
        if cell.ctype == XL_CELL_DATE:
            self.wb.sheets()
            if not self.tb:
                self.tb = copy(self.wb)
            curval = datetime(*xldate_as_tuple(cell.value, self.wb.datemode))
            newval = curval+timedelta(int(numdays))
            plain = easyxf('', num_format_str='d.M.yy')
            self.tb.get_sheet(0).write(int(row),
                                       int(column),
                                       newval,
                                       plain)

    def subtract_from_date(self, row, column, numdays):
        """Subtracts a number of days from the
        date in the indicated cell.
        """
        cell = self.wb.get_sheet(0).cell(int(row), int(column))
        if cell.ctype == XL_CELL_DATE:
            self.wb.sheets()
            if not self.tb:
                self.tb = copy(self.wb)
            curval = datetime(*xldate_as_tuple(cell.value, self.wb.datemode))
            newval = curval-timedelta(int(numdays))
            plain = easyxf('', num_format_str='d.M.yy')
            self.tb.get_sheet(0).write(int(row),
                                       int(column),
                                       newval,
                                       plain)

    def save_excel(self, fname):
        """Saves the Excel file indicated by fname"""
        print '*DEBUG* Got fname %s' % fname
        self.tb.save(os.path.join(self.slash, self.tmpDir, fname))

    def create_excel(self):
        """Creates a new Excel workbook"""
        self.tb = Workbook()
        self.tb.add_sheet('Sheet 1')
        
    def row_count(self,  sheetname):
        RowCount =self.wb.sheet_by_name(sheetname).nrows
        return  RowCount
        
    def col_count(self, sheetname):
        ColCount =self.wb.sheet_by_name(sheetname).ncols
        return  ColCount    
