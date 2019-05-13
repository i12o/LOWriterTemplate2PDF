#!/usr/bin/python3

import os
workfolder=os.path.dirname(__file__)
if workfolder:
    print(__file__)
    os.chdir(os.path.join(workfolder,'../'))
import sys
sys.path.append(os.curdir)
import argparse
import csv
import logging
import atexit
import re
import string

import pandas as pd

import InsertionProcess
import filters.for_hagaki
import InsertionProcess.common_options

# このディレクトリ下にチェックボックスなどの画像があるとする。
# テンプレートファイルからの相対でも、絶対パスでも良い
SYMBOLSDIRPATH = 'icons'
LOPORT="socket,host=localhost,port=2083;urp;"
PDFEXPORTDIR = "./pdfout"

optionvalues = {
    'loglevel' : logging.WARNING,
    'use_iconsdir' : SYMBOLSDIRPATH,
    'use_loport' : LOPORT,
    'use_pdfexport' : PDFEXPORTDIR,
    'template_file' :  None,
    'csv_file' : None,
    'spillfix' : None,
}

parser = InsertionProcess.common_options.args_setup(
    "Fill each record of data into Template, and export PDF",
    optionvalues
)

parser.add_argument('--no-page-select', action="store_true",
                    help="for testing: export all page in template")

args = InsertionProcess.common_options.args_parse(parser,optionvalues)

logging.basicConfig(level = optionvalues['loglevel'])

converters = [
    filters.for_hagaki.zip,
    filters.for_hagaki.one_char_name,
]

if not args.no_page_select:
    converters.append(filters.for_hagaki.select_one_style)

def filenamer(num,record,obj,*,pcr=None,pcindex=0):
    pdffilename = '{:04d}.pdf'.format(num+1)
    return pdffilename

try:
    inserter \
        = InsertionProcess.InsertionProcess(optionvalues['template_file'],
                                            optionvalues['use_loport'],
                                            symbolsdir=optionvalues['use_iconsdir'],
                                            pdfoutdir=optionvalues['use_pdfexport'],
                                            spillfix=optionvalues['spillfix'],
                                            filenamer=filenamer,
                                            data_converter=converters
        )

except:
    inserter.cleanup()
    logging.error("Can't connect LibreOffice, or open Document {}" \
                  .format(optionvalues['template_file']))
    exit(1)

atexit.register(inserter.cleanup)
logging.info("Use template {}, load icons from {} and export PDF to {}" \
             .format(optionvalues['template_file'],
                     optionvalues['use_iconsdir'],
                     optionvalues['use_pdfexport']))

# pandas を使って CSV を読み込む。
# read_table はタブ切りを読み込み、dtype=str と指定することで
# 全てのレコードを文字列として読み込む
csvrecords = pd.read_table(optionvalues['csv_file'],dtype=str)
# 全カラムが空値の行を削除する
csvrecords = csvrecords.dropna(how='all')
# NaN を空文字で置き換えておく
csvrecords = csvrecords.fillna('')
logging.info("Load csv{} done".format(optionvalues['csv_file']))

try:
    inserter.main_loop_pandas(csvrecords)

finally:
    pass

logging.info("Done.")
