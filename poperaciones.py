#!/usr/bin/env python2
# -*- coding: utf-8-*-

"""
    poperaciones
    ~~~~~~~~~~~~

    Lista la cantidad de operaciones realizadas por cada vendedor agrupadas
    mensualmente.
"""

from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
from dateutil.parser import parse as date_parse
import xlwt

from sqlalchemy.sql import func, extract, desc, asc

from nobix.config import load_config
from nobix.db import setup_db, Session
from nobix.models import Documento

ezxf = xlwt.easyxf
xf_map = {
    'heading': ezxf('font: bold on; align: vert centre, horiz left'),
    'text': ezxf(),
    'month': ezxf('font: bold on; align: vert centre, horiz center', num_format_str='mm-yyyy'),
    'number': ezxf(num_format_str='0')
}

config = load_config()
doctypes = [u'FAC', u'FAA', u'REM']

def init_nobix_db():
    setup_db(config.database_uri)
    session = Session()
    return session

def dump_sheet(sheet, options):
    session = init_nobix_db()

    months = int(options.get('months', None))
    upto = options.get('upto', None)

    dates = list(reversed([(upto - relativedelta(months=i)) for i in xrange(months)]))

    vend_names = {}
    for c in config.vendedores.keys():
        vname = config.vendedores.get(c, {}).get('nombre', None)
        vend_names.setdefault(vname, list()).append(c)

    base_query = Documento.query.filter(Documento.tipo.in_(doctypes))

    heads = [d.strftime("%m-%Y") for d in dates]

    row = 0
    sheet.write(row, 0, 'vendedor', xf_map['heading'])
    for c, h in enumerate(heads):
        sheet.write(row, c+1, h, xf_map['heading'])

    for vend in sorted(vend_names.keys()):
        row += 1
        col = 0
        vcodes = vend_names.get(vend)
        query = base_query.filter(Documento.vendedor.in_(vcodes))
        sheet.write(row, col, vend, xf_map['text'])
        for d in dates:
            col += 1
            q = query.filter(extract('year', Documento.fecha)==d.year)\
                     .filter(extract('month', Documento.fecha)==d.month)
            sheet.write(row, col, q.count(), xf_map['number'])


def write_xls(options):
    filename = options.pop('outfile')
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Cantidad Operaciones')
    dump_sheet(ws, options)
    wb.save(filename)
    print "Saved output to: %s" % filename

if __name__ == '__main__':
    from optparse import OptionParser

    parser = OptionParser()
    parser.add_option("-m", "--months", dest="months", default=6,
                      help="Cantidad de meses a mostrar [default: 6]")
    parser.add_option("-u", "--upto", dest="upto", default=None,
                      help="Operaciones hasta mm-yyyy [default: this month]")
    parser.add_option("-o", "--outfile", dest="outfile",
                      default="hist_operaciones.xls",
                      help="Archivo de salida [opcional]")

    (options, args) = parser.parse_args()

    if options.upto:
        #datep = dparser(dayfirst=True)
        options.upto = date_parse('01-' + options.upto, dayfirst=True)
    else:
        options.upto = date.today()

    o = dict(options.__dict__)
    write_xls(o)
