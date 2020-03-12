# -*- coding:utf-8 -*-

import codecs
import datetime

import pytz
import screen
import xlwt
from django.conf import settings
from django.utils import timezone

from .compat import basestring, str


@property
def as_xls(self):
    book = xlwt.Workbook(encoding=self.encoding)
    sheet = book.add_sheet(self.sheet_name)

    styles = {
        'datetime': xlwt.easyxf(num_format_str='yyyy-mm-dd hh:mm:ss'),
        'date': xlwt.easyxf(num_format_str='yyyy-mm-dd'),
        'time': xlwt.easyxf(num_format_str='hh:mm:ss'),
        'font': xlwt.easyxf('%s %s' % ('font:', self.font)),
        'default': xlwt.Style.default_style,
    }

    widths = {}
    for rowx, row in enumerate(self.data):
        for colx, value in enumerate(row):
            if value is None and self.blanks_for_none:
                value = ''

            if isinstance(value, datetime.datetime):
                if timezone.is_aware(value):
                    value = timezone.make_naive(value, pytz.timezone(settings.TIME_ZONE))
                cell_style = styles['datetime']
            elif isinstance(value, datetime.date):
                cell_style = styles['date']
            elif isinstance(value, datetime.time):
                cell_style = styles['time']
            elif self.font:
                cell_style = styles['font']
            else:
                cell_style = styles['default']

            sheet.write(rowx, colx, value, style=cell_style)

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
            if self.auto_adjust_width:
                width = screen.calc_width(value) * 256 if isinstance(value, basestring) else screen.calc_width(str(value)) * 256
                if width > widths.get(colx, 0):
                    width = min(width, self.EXCEL_MAXIMUM_ALLOWED_COLUMN_WIDTH)
                    widths[colx] = width
                    sheet.col(colx).width = width

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
