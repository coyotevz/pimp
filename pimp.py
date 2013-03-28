# -*- coding: utf-8 -*-

"""
    pimp
    ~~~~

    xls price importer
"""

from os import path
from collections import defaultdict
from decimal import Decimal
from xlrd import open_workbook
from xlrd import XL_CELL_TEXT, XL_CELL_NUMBER
from xlutils.copy import copy as copy_workbook

#from nobix.db import setup_db, Session
#from nobix.models import Articulo
#from nobix.config import load_config

q = Decimal('0.01')

test_file = 'test_file/griferia.xls'


def init_nobix_db():
    config = load_config()
    setup_db(config.database_uri)
    session = Session()
    return session


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


def cast(cell):
    if cell.ctype is XL_CELL_NUMBER:
        if isinstance(cell.value, float):
            return Decimal(str(cell.value)).quantize(q)
        return int(cell.value)
    elif cell.ctype is XL_CELL_TEXT:
        return cell.value.strip()
    return cell.value


def process_sheet(sheet, spec):
    ref_name = spec['ref'].keys()[0]
    rev_col = spec['ref'][ref_name]
    reads = list(spec['read'].iteritems())
    updates = list(spec['update'].iteritems())

    if 'status' in spec:
        def log_status(msg, row):
            pass
    else:
        def log_status(msg, row):
            pass

    def ignore(cell):
        return cell.ctype == XL_CELL_TEXT and cell.value.startswith(u'!')
    for r in range(spec['startrow'], sheet.nrows):
        row = sheet.row(r)
        if any(map(ignore, row)):
            continue
        ref_val = cast(row[rev_col])
        if ref_val:
# 1st read fields
# 2nd update fields
            newval = dict([(k, cast(row[i])) for k, i in updates])
# 3rd stamp status
            print ref_val, newval

if __name__ == '__main__':
    workbook = open_workbook(test_file, on_demand=True)
    test_sheet = workbook.sheet_by_name(u'GRIFERIA')
    spec = get_spec(test_sheet)
    process_sheet(test_sheet, spec)
