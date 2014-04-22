#!/usr/bin/env python2
# -*- coding: utf-8 -*-

"""
    top200
    ~~~~~~

    Entrega un .xls con los 200 articulos más vendidos de los últimos 12 meses.
"""

import sys
from decimal import Decimal
from datetime import date
import xlwt

from sqlalchemy.sql import func, desc, asc

from nobix.config import load_config
from nobix.db import setup_db, Session
from nobix.models import Articulo, ItemDocumento, Documento

today = date.today()
end = today.replace(month=today.month-1)
start = end.replace(year=end.year-1)

ezxf = xlwt.easyxf
xf_map = {
    'heading': ezxf('font: bold on; align: horiz left'),
    'date': ezxf(num_format_str='yyy-mm-dd'),
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

    query = session.query(Articulo, func.sum(ItemDocumento.cantidad).label("custom_tot"))\
                .join(ItemDocumento, ItemDocumento.articulo_id==Articulo.id)\
                .join(Documento, ItemDocumento.documento_id==Documento.id)\
                .filter(Documento.tipo.in_(sal))\
                .filter(Documento.fecha.between(start, end))\
                .group_by(Articulo)

    q1 = query.order_by(desc("custom_tot"))
    q2 = query.order_by(asc("custom_tot"))

    heads = [u'Código', u'Descripción', u'Agrupación']

    row = 0
    for c, h in enumerate(heads):
        sheet.write(row, c, h, xf_map['heading'])

    row += 1
    for art, q in q1[:200]:
        sheet.write(row, 0, art.codigo, xf_map['text'])
        sheet.write(row, 1, art.descripcion, xf_map['text'])
        sheet.write(row, 2, art.agrupacion, xf_map['text'])
        sheet.write(row, 3, q, xf_map['text'])
        row += 1

    sheet.write(row, 0, "Lo menos vendido", xf_map['text'])
    row += 1

    for c, h in enumerate(heads):
        sheet.write(row, c, h, xf_map['heading'])

    row += 1
    for art, q in q2[:200]:
        sheet.write(row, 0, art.codigo, xf_map['text'])
        sheet.write(row, 1, art.descripcion, xf_map['text'])
        sheet.write(row, 2, art.agrupacion, xf_map['text'])
        sheet.write(row, 3, q, xf_map['text'])
        row += 1

def write_xls(options):
    filename = options.pop('outfile')
    wb = xlwt.Workbook()
    ws = wb.add_sheet('top200')

    dump_sheet(ws, options)

    wb.save(filename)


if __name__ == '__main__':

    from optparse import OptionParser

    parser = OptionParser()
    parser.add_option("-o", "--outfile", dest="outfile", default="top200.xls",
                      help="Archivo de salida [opcional]")

    (options, args) = parser.parse_args()

    o = dict(options.__dict__)

    write_xls(o)
