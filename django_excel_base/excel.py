# -*- coding:utf-8 -*-

import codecs
import datetime
import json
from collections import deque

import pytz
import screen
import xlwt
from django.conf import settings
from django.utils import timezone

from .compat import basestring, str


def get_merged_rows_num(data):
    return max(sum(get_merged_rows_num(d) if isinstance(d, list) else 0 for d in data), 1)


def row_preprocessing(data, mapping):
    field_key = mapping.get('field_key')
    field_key = field_key if isinstance(field_key, list) else [field_key]
    data_key = mapping.get('data_key')
    nextmapping = mapping.get('next')
    field_value = [data.get(key) for key in field_key]
    return field_value + [data_preprocessing(data.get(data_key), nextmapping)] if nextmapping else field_value


def data_preprocessing(data, mapping):
    data = data if isinstance(data, list) else [data]
    return [row_preprocessing(item, mapping) for item in data]


def get_cell_styles(self):
    al = xlwt.Alignment()
    al.horz = self.horz
    al.vert = self.vert

    datetime_style = xlwt.easyxf(num_format_str='yyyy-mm-dd hh:mm:ss')
    datetime_style.alignment = al

    date_style = xlwt.easyxf(num_format_str='yyyy-mm-dd')
    date_style.alignment = al

    time_style = xlwt.easyxf(num_format_str='hh:mm:ss')
    time_style.alignment = al

    font_style = xlwt.easyxf('%s %s' % ('font:', self.font))
    font_style.alignment = al

    dafault_style = xlwt.Style.default_style
    dafault_style.alignment = al

    return {
        'datetime': datetime_style,
        'date': date_style,
        'time': time_style,
        'font': font_style,
        'default': dafault_style,
    }


def get_cell_info(self, value, cell_styles):
    if value is None and self.blanks_for_none:
        value = ''

    if isinstance(value, datetime.datetime):
        if timezone.is_aware(value):
            value = timezone.make_naive(value, pytz.timezone(settings.TIME_ZONE))
        cell_style = cell_styles['datetime']
    elif isinstance(value, datetime.date):
        cell_style = cell_styles['date']
    elif isinstance(value, datetime.time):
        cell_style = cell_styles['time']
    elif self.font:
        cell_style = cell_styles['font']
    else:
        cell_style = cell_styles['default']

    return value, cell_style


def auto_adjust_width(self, sheet, colx, value, widths):
    # Columns have a property for setting the width.
    # The value is an integer specifying the size measured in 1/256
    # of the width of the character '0' as it appears in the sheet's default font.
    # xlwt creates columns with a default width of 2962, roughly equivalent to 11 characters wide.
    #
    # https://github.com/python-excel/xlwt/blob/master/xlwt/BIFFRecords.py#L1675
    # Offset  Size    Contents
    # 4       2       Width of the columns in 1/256 of the width of the zero character, using default font
    #                 (first FONT record in the file)
    #
    # Default Width: https://github.com/python-excel/xlwt/blob/master/xlwt/Column.py#L14
    # self.width = 0x0B92
    if not self.auto_adjust_width:
        return
    width = screen.calc_width(value) * 256 if isinstance(value, basestring) else screen.calc_width(str(value)) * 256
    if width <= widths.get(colx, 0):
        return
    width = min(width, self.EXCEL_MAXIMUM_ALLOWED_COLUMN_WIDTH)
    widths[colx] = width
    sheet.col(colx).width = max(width, self.min_cell_width)


@property
def as_xls(self):
    if not isinstance(self.data, dict):
        self.data = {self.sheet_name: self.data}

    cell_styles = get_cell_styles(self)

    book = xlwt.Workbook(encoding=self.encoding)

    for sheet_name, sheet_data in self.data.items():
        sheet = book.add_sheet(sheet_name)

        widths = {}
        for rowx, row in enumerate(sheet_data):
            for colx, value in enumerate(row):
                if value is None and self.blanks_for_none:
                    value = ''
                value, cell_style = get_cell_info(self, value, cell_styles)
                sheet.write(rowx, colx, value, style=cell_style)
                auto_adjust_width(self, sheet, colx, value, widths)

    book.save(self.output)


@property
def as_row_merge_xls(self):
    if not isinstance(self.data, dict):
        self.data = {self.sheet_name: self.data}

    cell_styles = get_cell_styles(self)

    book = xlwt.Workbook(encoding=self.encoding)

    for sheet_name, sheet_data in self.data.items():
        sheet = book.add_sheet(sheet_name)

        widths = {}
        rowx = 0  # 行起始索引
        for _, row in enumerate(sheet_data):
            # Max row number for current row
            rowmax = max([(len(r) if isinstance(r, list) else 1) for r in row])
            for colx, value in enumerate(row):
                if isinstance(value, list):
                    for vx, val in enumerate(value):
                        val, cell_style = get_cell_info(self, val, cell_styles)
                        sheet.write(rowx + vx, colx, val, style=cell_style)
                        auto_adjust_width(self, sheet, colx, val, widths)
                else:
                    value, cell_style = get_cell_info(self, value, cell_styles)
                    sheet.write_merge(rowx, rowx + rowmax - 1, colx, colx, value, style=cell_style)
                    auto_adjust_width(self, sheet, colx, value, widths)

            rowx += rowmax  # 更新行起始索引

    book.save(self.output)


@property
def as_list_row_merge_xls(self):
    if not isinstance(self.data, dict):
        self.data = {self.sheet_name: self.data}

    cell_styles = get_cell_styles(self)

    book = xlwt.Workbook(encoding=self.encoding)

    q = deque()
    for sheet_name, sheet_data in self.data.items():
        sheet = book.add_sheet(sheet_name)

        widths = {}
        rowx = 0  # 行起始索引
        colx = 0
        for _, row in enumerate(sheet_data):
            rowmax = 1
            _rowx = rowx
            _rowxd = {}
            q.append(json.dumps([colx, row], ensure_ascii=False))
            while q:
                colx, rowdata = json.loads(q.popleft())
                mergedrowsnum = get_merged_rows_num(rowdata)
                rowmax = max(rowmax, mergedrowsnum)
                for idx, value in enumerate(rowdata):
                    _colx = colx + idx
                    if isinstance(value, list):
                        for d in value:
                            q.append(json.dumps([_colx, d], ensure_ascii=False))
                    else:
                        _rowx = _rowxd.get(_colx, rowx)
                        _rowxd[_colx] = _rowx + mergedrowsnum
                        value, cell_style = get_cell_info(self, value, cell_styles)
                        sheet.write_merge(_rowx, _rowx + mergedrowsnum - 1, _colx, _colx, value, style=cell_style)
                        auto_adjust_width(self, sheet, _colx, value, widths)

            rowx += rowmax  # 更新行起始索引
            colx = 0

    book.save(self.output)


@property
def as_dict_row_merge_xls(self):
    if not isinstance(self.data, dict):
        self.data = {self.sheet_name: {'data': self.data, 'mapping': self.mapping, 'headers': self.headers}}

    cell_styles = get_cell_styles(self)

    book = xlwt.Workbook(encoding=self.encoding)

    q = deque()
    for sheet_name, sheet_data in self.data.items():
        data = sheet_data.get('data')
        mapping = sheet_data.get('mapping')
        headers = sheet_data.get('headers')
        sheet_data = ([headers] if headers else []) + data_preprocessing(data, mapping)

        sheet = book.add_sheet(sheet_name)

        widths = {}
        rowx = 0  # 行起始索引
        colx = 0
        for _, row in enumerate(sheet_data):
            rowmax = 1
            _rowx = rowx
            _rowxd = {}
            q.append(json.dumps([colx, row], ensure_ascii=False))
            while q:
                colx, rowdata = json.loads(q.popleft())
                mergedrowsnum = get_merged_rows_num(rowdata)
                rowmax = max(rowmax, mergedrowsnum)
                for idx, value in enumerate(rowdata):
                    _colx = colx + idx
                    if isinstance(value, list):
                        for d in value:
                            q.append(json.dumps([_colx, d], ensure_ascii=False))
                    else:
                        _rowx = _rowxd.get(_colx, rowx)
                        _rowxd[_colx] = _rowx + mergedrowsnum
                        value, cell_style = get_cell_info(self, value, cell_styles)
                        sheet.write_merge(_rowx, _rowx + mergedrowsnum - 1, _colx, _colx, value, style=cell_style)
                        auto_adjust_width(self, sheet, _colx, value, widths)

            rowx += rowmax  # 更新行起始索引
            colx = 0

    book.save(self.output)


@property
def as_csv(self):
    # https://stackoverflow.com/questions/4348802/how-can-i-output-a-utf-8-csv-in-php-that-excel-will-read-properly?answertab=votes
    # https://stackoverflow.com/questions/155097/microsoft-excel-mangles-diacritics-in-csv-files/1648671#1648671
    # https://wiki.scn.sap.com/wiki/display/ABAP/CSV+tests+of+encoding+and+column+separator?original_fqdn=wiki.sdn.sap.com
    if self.encoding == 'utf-8-sig':
        self.output.write(codecs.BOM_UTF8)
    for row in self.data:
        out_row = []
        for value in row:
            if value is None and self.blanks_for_none:
                value = ''
            if not isinstance(value, basestring):
                value = str(value)
            out_row.append(value.replace('"', '""').encode(self.encoding))
        self.output.write(b'"%s"\n' % b'","'.join(out_row))
