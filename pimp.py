# -*- coding: utf-8 -*-

"""
    pimp
    ~~~~

    xls price importer
"""

from os import path
from collections import defaultdict
from xlrd import open_workbook, XL_CELL_TEXT
from xlutils.copy import copy as copy_workbook

test_file = 'test_file/griferia.xls'


def get_spec(sheet):
    ref_found = False
    spec = defaultdict(dict)
    for n in range(sheet.nrows):
        for i, v in enumerate(sheet.row_values(n)):
            if isinstance(v, basestring):
                if v.startswith(u'#'):
                    spec['ref'][v[1:]] = i
                    spec['startrow'] = n + 1
                    ref_found = True
                elif v.startswith(u'$'):
                    spec['update'][v[1:]] = i
                elif v == '@status':
                    spec['status'] = i
                elif v.startswith(u'@'):
                    spec['read'][v[1:]] = i
        if ref_found:
            break
    if not ref_found:
        raise Exception("There is no reference in sheet '%s'" % sheet.name)
    return dict(spec)


def process_sheet(sheet, spec):
    def ignore(cell):
        return cell.ctype == XL_CELL_TEXT and cell.value.startswith(u'!')
    for r in range(spec['startrow'], sheet.nrows):
        if any(map(ignore, sheet.row(r))):
            continue


if __name__ == '__main__':
    workbook = open_workbook(test_file, on_demand=True)
    test_sheet = workbook.sheet_by_name(u'GRIFERIA')
    spec = get_spec(test_sheet)
    process_sheet(test_sheet, spec)
