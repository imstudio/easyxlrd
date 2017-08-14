# coding: utf-8

import os
import xlrd


class SheetData(object):
    def __init__(self, sheet_data):
        self._data = sheet_data
        self._itr_pos = 0

    def __getitem__(self, row_index):
        if row_index >= self._data.nrows:
            raise KeyError('invalid row index:{}'.format(row_index))
        return [self._data.cell(row_index, j).value for j in xrange(self._data.ncols)]

    def __iter__(self):
        self._itr_pos = 0
        return self

    def next(self):
        if self._itr_pos < self._data.nrows:
            r = [self._data.cell(self._itr_pos, j).value for j in xrange(self._data.ncols)]
            self._itr_pos += 1
            return r
        raise StopIteration()


class EasyXlrd(object):
    def __init__(self, file_path):
        if not os.path.exists(file_path):
            raise IOError('file {} not exist.'.format(file_path))
        if not os.path.isfile(file_path):
            raise IOError('{} is dir.'.format(file_path))
        self._xls = xlrd.open_workbook(file_path)
        self._itr_pos = 0

    def __getitem__(self, sheet_id):
        if sheet_id >= self._xls.nsheets:
            raise KeyError('invalid sheet id: {}'.format(sheet_id))
        return SheetData(self._xls.sheet_by_index(sheet_id))

    def __iter__(self):
        self._itr_pos = 0
        return self

    def next(self):
        if self._itr_pos < self._xls.nsheets:
            r = SheetData(self._xls.sheet_by_index(self._itr_pos))
            self._itr_pos += 1
            return r
        raise StopIteration()


if __name__ == '__main__':
    import sys


    def do_test():
        file_name = sys.argv[1]
        ex = EasyXlrd(file_name)
        for sheet_data in ex:
            for row_data in sheet_data:
                print row_data


    do_test()
