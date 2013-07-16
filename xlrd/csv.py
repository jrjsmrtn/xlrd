# -*- coding: utf-8; mode: python; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- vim:fenc=utf-8:ft=py:et:sw=4:ts=4:sts=4

# Copyright (c) 2013 Georges Martin under a BSD licence

import csv
from . import sheet

__all__ = (
    'SheetReader',
    'SheetDictReader',
    )


class SheetReader(object):
    """A csv.reader-compatible reader for xlrd.sheet."""

    def __init__(self, sheet, *args, **kw):
        if not isinstance(sheet, xlrd.sheet.Sheet):
            raise TypeError()
        self.sheet = sheet
        self._nrows = sheet.nrows
        self.line_num = 0
    
    def __iter__(self):
        return self
    
    def next(self):
        self.line_num += 1
        if self.line_num >= self._nrows:
            raise StopIteration
        else:
            return self.sheet.row_values(self.line_num)


class SheetDictReader(csv.DictReader):
    """A csv.DictReader-compatible DictReader for xlrd.sheet."""

    def __init__(self, sheet, fieldnames=None, restkey=None, restval=None, *args, **kw):
        if not isinstance(sheet, xlrd.sheet.Sheet):
            raise TypeError()
        self.sheet = sheet
        self.reader = SheetReader(sheet, *args, **kw)
        self._fieldnames = fieldnames
        self.restkey = restkey
        self.restval = restval
        self.line_num = 0
