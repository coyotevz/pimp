#!/usr/bin/env python2
# -*- coding: utf-8 -*-

"""
    pmov
    ~~~~

    Lista los movimientos de articulos por agrupación y fecha.
"""

from decimal import Decimal
from datetime import date
import xlwt

from sqlalchemy.sql import func, extract, desc, asc

from nobix.config import load_config
from nobix.db import setup_db, Session
from nobix.models import Articulo, Documento, ItemDocumento

today = date.today()

ezxf = xlwt.easyxf
xf_map = {
}

config = load_config()

def init_nobix_db():
    setup_db(config.database_uri)
    session = Session()
    return session

def dump_sheet(sheet, options):
    session = init_nobix_db()

    since = options.get('since', None)
    upto = options.get('upto', None)
    groups = options.get('groups', None)

    query = Articulo.query
    if groups:
        query = query.filter(Articulo.agrupacion.in_(groups))

    query = query.order_by(Articulo.agrupacion, Articulo.codigo)

    heads = ['codigo', 'descripcion', 'agrupacion', 'entradas', 'salidas']

    row = 0
    for c, h in enumerate(heads):
        sheet.write(row, c, h, xf_map['heading'])

    row += 1
    for item in query:
        sheet.write(row, 0, item.codigo, xf_map['text'])
        sheet.write(row, 1, item.descripcion, xf_map['text'])
        sheet.write(row, 2, item.agrupacion, xf_map['text'])
        sheet.write(row, 3, item.entradas, xf_map['number'])
        sheet.write(row, 4, item.salidas, xf_map['number'])
        row += 1

def write_xls(options):
    filename = options.pop('outfile')
    wb = xlwt.Workbook()
    if options['groups'] is None:
        ws = wb.add_sheet('Movimientos')
        dump_sheet(ws, options)
    else:
        for g in options['groups']:
            o = options.copy()
            o['groups'] = g
            ws = wb.add_sheet(g)
            dump_sheet(ws, o)
    wb.save(filename)

if __name__ == '__main__':
    from optparse import OptionParser

    parser = OptionParser()
    parser.add_option("-g", "--groups", dest="groups", default="*",
                      help="Agrupaciones (separadas por coma) [opcional]")
    parser.add_option("-s", "--since", dest="since", default=None,
                      help="Movimientos desde [opcional]")
    parser.add_option("-u", "--upto", dest="upto", default=None,
                      help="Movimientos hasta [opcional]")
    parser.add_option("-o", "--outfile", dest="outfile", default="movimientos.xls",
                      help="Archivo de salida [opcional]")

    (options, args) = parser.parse_args()

    if options.groups:
        if options.groups == '*':
            options.groups = None
        else:
            options.groups = [unicode(g) for g in options.groups.split(",")]

    if options.since:
        options.since = datetime.strptime(options.since, "%Y-%m-%d")

    if options.upto:
        options.upto = datetime.strptime(options.upto, "%Y-%m-%d")

    o = dict(options.__dict__)
    write_xls(o)
