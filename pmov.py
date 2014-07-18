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
    pass

def write_xls(options):
    pass

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
