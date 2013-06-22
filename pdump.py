# -*- coding: utf-8 -*-

"""
    pdump
    ~~~~~

    xls exporter for nobix
"""

import sys
from decimal import Decimal
from datetime import datetime
import xlwt

from nobix.config import load_config
from nobix.db import setup_db, Session
from nobix.models import Articulo

q = Decimal('0.01')
now = datetime.now()

ezxf = xlwt.easyxf
xf_map = {
    'heading': ezxf('font: bold on; align: wrap on, vert centre, horiz right'),
    'date': ezxf(num_format_str='yyyy-mm-dd'),
    'price': ezxf(num_format_str='0.00'),
    'text': ezxf(),
}

def init_nobix_db():
    config = load_config()
    setup_db(config.database_uri)
    session = Session()
    return session

def dump_sheet(sheet, options):

    session = init_nobix_db()

    since = options['since']
    upto = options.get('upto', None)
    groups = options.get('groups', None)

    query = Articulo.query.filter(Articulo.vigencia>=since)
    if upto:
        query = query.filter(Articulo.vigencia<=upto)
    if groups:
        query = query.filter(Articulo.agrupacion.in_(groups))

    query = query.order_by(Articulo.agrupacion, Articulo.codigo)

    heads = ['codigo', 'descripcion', 'agrupacion', 'proveedor', 'vigencia',
             'precio']

    row = 0
    for c, h in enumerate(heads):
        sheet.write(row, c, h, xf_map['heading'])

    row += 1
    for art in query:
        sheet.write(row, 0, art.codigo, xf_map['text'])
        sheet.write(row, 1, art.descripcion, xf_map['text'])
        sheet.write(row, 2, art.agrupacion, xf_map['text'])
        sheet.write(row, 3, art.proveedor, xf_map['text'])
        sheet.write(row, 4, art.vigencia, xf_map['date'])
        sheet.write(row, 5, art.precio, xf_map['price'])
        row += 1


def write_xls(options):

    filename = options.pop('outfile')
    wb = xlwt.Workbook()
    ws = wb.add_sheet('NOBIX_UPDATE')

    dump_sheet(ws, options)

    wb.save(filename)


if __name__ == '__main__':

    from optparse import OptionParser

    parser = OptionParser()
    parser.add_option("-s", "--since", dest="since",
                      help="Vigencia desde")
    parser.add_option("-u", "--upto", dest="upto", default=None,
                      help="Vigencia hasta [opcional]")
    parser.add_option("-g", "--groups", dest="groups", default="*",
                      help="Agrupaciones (separados por coma) [opcional]")
    parser.add_option("-o", "--outfile", dest="outfile", default="dump.xls",
                      help="Archivo de salida [opcional]")

    (options, args) = parser.parse_args()

    if not options.since:
        sys.exit("ERROR: You must provide --since argument at least. -h for help")

    if options.groups:
        if options.groups == '*':
            options.groups = None
        else:
            options.groups = options.groups.split(",")

    options.since = datetime.strptime(options.since, "%Y-%m-%d")

    if options.upto:
        options.upto = datetime.strptime(options.upto, "%Y-%m-%d")

    o = dict(options.__dict__)

    write_xls(o)
