#!/usr/bin/env python2
# -*- coding: utf-8 -*-

"""
    top200
    ~~~~~~

    Entrega un .xls con los 200 articulos más vendidos de los últimos 12 meses.
"""

import sys
from decimal import Decimal
from datetime import date, timedelta
import xlwt

from sqlalchemy.sql import func, extract, desc, asc

from nobix.config import load_config
from nobix.db import setup_db, Session
from nobix.models import Articulo, ItemDocumento, Documento

today = date.today()
end = today.replace(month=today.month-1, day=1)
start = end.replace(year=end.year-1)

date_range = [start + (n*timedelta(31)) for n in xrange(1, 13)]

ezxf = xlwt.easyxf
xf_map = {
    'heading': ezxf('font: bold on; align: horiz left'),
    'date': ezxf('font: bold on; align: horiz right', num_format_str='mm-yyyy'),
    'quantity': ezxf(num_format_str='0.00'),
    'text': ezxf(),
}

config = load_config()

def init_nobix_db():
    setup_db(config.database_uri)
    session = Session()
    return session

def dump_sheet(sheet, options):

    session = init_nobix_db()

    sal = [k for k, v in config.documentos.iteritems() if v['stock'] == u'salida']

    if options.get('price', False):
        f = func.sum(ItemDocumento.cantidad * ItemDocumento.precio)
    else:
        f = func.sum(ItemDocumento.cantidad)

    query = session.query(Articulo, f.label("custom_tot"))\
                .join((ItemDocumento, ItemDocumento.articulo_id==Articulo.id))\
                .join((Documento, ItemDocumento.documento_id==Documento.id))\
                .filter(Documento.fecha.between(start, end))\
                .filter(Documento.tipo.in_(sal))\
                .group_by(Articulo)

    q1 = query.order_by(desc("custom_tot"))
    q2 = query.order_by(asc("custom_tot"))

    heads = [u'Código', u'Descripción', u'Agrupación']

    row = 0
    for c, h in enumerate(heads):
        sheet.write(row, c, h, xf_map['heading'])
    for i, h in enumerate(date_range):
        sheet.write(row, i+c+1, h, xf_map['date'])

    row += 1
    for art, q in q1[:200]:
        sheet.write(row, 0, art.codigo, xf_map['text'])
        sheet.write(row, 1, art.descripcion, xf_map['text'])
        sheet.write(row, 2, art.agrupacion, xf_map['text'])
        col = 3
        quan = session.query(f)\
                .join((Documento, ItemDocumento.documento_id==Documento.id))\
                .filter(Documento.tipo.in_(sal))\
                .filter(ItemDocumento.articulo_id==art.id)
        for d in date_range:
            q = quan.filter(extract('year', Documento.fecha)==d.year)\
                    .filter(extract('month', Documento.fecha)==d.month)
            sheet.write(row, col, q.scalar(), xf_map['quantity'])
            col += 1
        row += 1

    #sheet.write(row, 0, "Lo menos vendido", xf_map['text'])
    #row += 1

    #for c, h in enumerate(heads):
    #    sheet.write(row, c, h, xf_map['heading'])

    #row += 1
    #for art, q in q2[:200]:
    #    sheet.write(row, 0, art.codigo, xf_map['text'])
    #    sheet.write(row, 1, art.descripcion, xf_map['text'])
    #    sheet.write(row, 2, art.agrupacion, xf_map['text'])
    #    sheet.write(row, 3, q, xf_map['text'])
    #    row += 1

def write_xls(options):
    filename = options.pop('outfile')
    wb = xlwt.Workbook()
    ws = wb.add_sheet('top200')

    dump_sheet(ws, options)

    wb.save(filename)


if __name__ == '__main__':

    from optparse import OptionParser

    parser = OptionParser()
    parser.add_option("-p", "--price", dest="price", default=False,
                      help="Analizar precios de venta")
    parser.add_option("-o", "--outfile", dest="outfile", default="top200.xls",
                      help="Archivo de salida [opcional]")

    (options, args) = parser.parse_args()

    o = dict(options.__dict__)

    write_xls(o)
