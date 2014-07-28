#!/usr/bin/env python2
# -*- coding: utf-8 -*-

"""
    dump_env
    ~~~~~~~~

    Exporta documentos tipo ENV (Envío)
"""

import sys
from decimal import Decimal
from datetime import date
import xlwt

from nobix.config import load_config
from nobix.db import setup_db, Session
from nobix.models import Documento

q = Decimal('0.01')
start = date(2014, 1, 1)

ezxf = xlwt.easyxf
xf_map = {
    'heading_bold': ezxf('font: bold on; border: bottom thin'),
    'heading_bold': ezxf('font: bold on; border: bottom thin; align: horiz left'),
    'heading_text': ezxf('font: bold on; border: bottom thin'),
    'heading_label': ezxf('font: bold on; border: bottom thin; align: horiz right'),
    'heading_date': ezxf('font: bold on; border: bottom thin', num_format_str='dd/mm/yyyy'),
    'heading_price': ezxf('font: bold on; border: bottom thin', num_format_str='#,##0.00'),
    'bold': ezxf('font: bold on; align: horiz left'),
    'date': ezxf(num_format_str='dd/mm/yyyy'),
    'int': ezxf(num_format_str='0'),
    'price': ezxf(num_format_str='#,##0.00'),
    'text': ezxf(),
}

def init_nobix_db():
    config = load_config()
    setup_db(config.database_uri)
    session = Session()
    return session

def dump_sheet(sheet, options):

    session = init_nobix_db()

    query = Documento.query.filter(Documento.tipo==u'ENV')\
                           .filter(Documento.fecha>=start)\
                           .order_by(Documento.fecha.asc())

    row = 0
    for doc in query:
        sheet.write(row, 0, u'Envío %d' % doc.numero, xf_map['heading_bold'])
        sheet.write(row, 1, doc.fecha, xf_map['heading_date'])
        sheet.write(row, 2, '', xf_map['heading_text'])
        sheet.write(row, 3, 'Total:', xf_map['heading_label'])
        sheet.write(row, 4, doc.total, xf_map['heading_price'])
        row += 1
        for item in doc.items:
            sheet.write(row, 1, item.codigo, xf_map['text'])
            sheet.write(row, 2, item.descripcion, xf_map['text'])
            sheet.write(row, 3, item.cantidad, xf_map['int'])
            sheet.write(row, 4, item.precio, xf_map['price'])
            row += 1
        row += 1


def write_xls(options):

    filename = options.pop('outfile')
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Envios')

    dump_sheet(ws, options)
    wb.save(filename)

if __name__ == '__main__':
    from optparse import OptionParser

    parser = OptionParser()
    parser.add_option("-o", "--outfile", dest="outfile", default="envios.xls",
                      help="Archivo de salida [opcional]")

    (options, args) = parser.parse_args()

    o = dict(options.__dict__)
    write_xls(o)
