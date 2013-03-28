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

from nobix.db import setup_db, Session
from nobix.models import Articulo
from nobix.config import load_config

from sqlalchemy.orm.exc import NoResultFound, MultipleResultsFound

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


def process_sheet(sheet, spec, outsheet, session):
    ref_name = spec['ref'].keys()[0]
    rev_col = spec['ref'][ref_name]
    reads = list(spec['read'].iteritems())
    updates = list(spec['update'].iteritems())

    if 'status' in spec:
        def log_status(msg, row):
            outsheet.write(row, spec['status'], msg)
    else:
        def log_status(msg, row):
            print "%s (#%s)" % (msg, row)

    def ignore(cell):
        return cell.ctype == XL_CELL_TEXT and cell.value.startswith(u'!')

    for r in range(spec['startrow'], sheet.nrows):
        row = sheet.row(r)
        if any(map(ignore, row)):
            continue

        ref_val = cast(row[rev_col])
        if ref_val:
            # 0st retrieve article by ref
            try:
                art = session.query(Articulo).filter(getattr(Articulo, ref_name) == ref_val).one()
            except NoResultFound:
                log_status("ERR: No se encuentra articulo que cumpla con '%s==%s'" % (ref_name, ref_val), r)
                continue
            except MultipleResultsFound:
                log_status("ERR: La condición '%s==%s' arroja multiples resultados" % (ref_name, rev_val), r)
                continue
            except Exception as e:
                log_status("EXCEPTION: %s" % " ".join(e.args), r)
                continue
            # 1st read fields
            # 2nd update fields
            toupdate = dict([(k, cast(row[i])) for k, i in updates])
            # 3rd stamp status
            print ref_val, toupdate

if __name__ == '__main__':
    sheet_name = u'GRIFERIA'
    workbook = open_workbook(test_file, on_demand=True, formatting_info=True)
    test_sheet = workbook.sheet_by_name(sheet_name)
    sheet_index = workbook.sheet_names().index(sheet_name)
    out_wb = copy_workbook(workbook)
    out_sheet = out_wb.get_sheet(sheet_index)
    spec = get_spec(test_sheet)
    print 'spec:', spec
    session = init_nobix_db()
    process_sheet(test_sheet, spec, out_sheet, session)
    out_wb.save('test_output.xls')
