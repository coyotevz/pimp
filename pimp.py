# -*- coding: utf-8 -*-

"""
    pimp
    ~~~~

    xls price importer
"""

import sys
from os import path
from collections import defaultdict
from decimal import Decimal
from datetime import datetime

from xlrd import open_workbook
from xlrd import XL_CELL_TEXT, XL_CELL_NUMBER
from xlutils.copy import copy as copy_workbook

from sqlalchemy.orm.exc import NoResultFound, MultipleResultsFound

from nobix.db import setup_db, Session
from nobix.models import Articulo
from nobix.config import load_config

q = Decimal('0.01')
now = datetime.now()


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
    return spec


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
    has_vigencia = 'vigencia' in spec['update']
    must_update_vigencia = (not has_vigencia) and ('precio' in spec['update'])
    has_status = 'status' in spec

    if has_status:
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
        #print "processing %s (%s)" % (ref_val, r),
        if ref_val:
            # 0st retrieve article by ref
            #print "OK"
            try:
                art = session.query(Articulo)\
                             .filter(getattr(Articulo, ref_name) == ref_val)\
                             .one()
            except NoResultFound:
                log_status("ERR: No se encuentra articulo que cumpla " +
                           "con '%s==%s'" % (ref_name, ref_val), r)
                continue
            except MultipleResultsFound:
                log_status("ERR: La condición '%s==%s' arroja multiples " +
                           "resultados" % (ref_name, rev_val), r)
                continue
            except Exception as e:
                log_status("EXCEPTION: %s" % " ".join(e.args), r)
                continue
            # 1st read fields
            for rkey, rcol in reads:
                val = getattr(art, rkey)
                outsheet.write(r, rcol, val)
            msg = "READ OK"
            # 2nd update fields
            toupdate = [(k, cast(row[i])) for k, i in updates]
            for ukey, uval in toupdate:
                setattr(art, ukey, uval)
            if must_update_vigencia:
                art.vigencia = now
            try:
                session.commit()
                if toupdate:
                    msg = "UPDATE OK"
            except Exception as e:
                session.rollback()
                msg = " ".join(e.args)
            # 3rd stamp status
            log_status(msg, r)
        else:
            #print "BAD"
            pass


def process_book(args=None):
    if args is None:
        args = sys.argv[1:]

    if len(args) == 0:
        sys.exit("Debe proveer el archivo de entrada *.xls")

    filename = args[0]
    if not path.exists(filename) and path.isfile(filename):
        sys.exit("El archivo '%s' no existe." % filename)

    fnparts = filename.rpartition('.')
    outfilename = fnparts[0] + '-out' + ''.join(fnparts[1:])

    workbook = open_workbook(filename, on_demand=True, formatting_info=True)
    out_wb = copy_workbook(workbook)

    sheet_names = workbook.sheet_names()

    session = init_nobix_db()

    for idx, name in enumerate(sheet_names):
        if name.lower().startswith(u'nobix_update'):
            input_sheet = workbook.sheet_by_index(idx)
            output_sheet = out_wb.get_sheet(idx)
            spec = get_spec(input_sheet)
            process_sheet(input_sheet, spec, output_sheet, session)
            print "processed %s" % name
    out_wb.save(outfilename)
    print "saved to %s" % outfilename


if __name__ == '__main__':
    process_book()
