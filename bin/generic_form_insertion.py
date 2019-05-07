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

# このディレクトリ下にチェックボックスなどの画像があるとする。
# テンプレートファイルからの相対でも、絶対パスでも良い
SYMBOLSDIRPATH = 'icons'
LOPORT="socket,host=localhost,port=2083;urp;"
PDFEXPORTDIR = "./pdfout"


parser = argparse.ArgumentParser(description= \
                                 "Fill each record of data into Template, and export PDF",
                                 formatter_class=argparse.RawTextHelpFormatter
)
dgroup = parser.add_mutually_exclusive_group()
dgroup.add_argument("-v", "--verbose",
                    default=0, action="count",
                    help="Increase verbosity"
)
dgroup.add_argument("-q", "--quiet",
                    action="store_true",
                    help="Run quietly"
)
parser.add_argument("--icons-path",
                    help="Icons(series of images) directory. Default relative to template's './icons'")
parser.add_argument("--pdf-export",
                    help="PDF export directory, Default './pdfout'")
parser.add_argument("--loport",
                    help="UNO-URL, to connect to LibreOffice\n  ex. {}".format(LOPORT))
parser.add_argument("-t", "--odt",
                    help=".odt file(LoWriter document) to be used as Template")
parser.add_argument("csv",
                    help="CSV records to be inserted into template")
args = parser.parse_args()

# # 以下は、ファイルをimportする場合や cronjob から実行する場合のためのもの。
# # .py ファイルがあるディレクトリにカレントディレクトリを移動する
# workfolder=os.path.dirname(__file__)
# if workfolder:
#     print(__file__)
#     os.chdir(workfolder)

loglevel = logging.WARNING
if args.quiet:
    loglevel = logging.ERROR
else:
    if args.verbose >= 2:
        loglevel = logging.DEBUG
    elif args.verbose >=1:
        loglevel = logging.INFO

logging.basicConfig(level = loglevel)

use_iconsdir = SYMBOLSDIRPATH
use_loport = LOPORT
use_pdfexport = PDFEXPORTDIR
template_file = None
csv_file = None

if args.icons_path:
    use_iconsdir = args.icons_path
if args.pdf_export:
    use_pdfexport = args.pdf_export
if args.loport:
    use_loport = args.LOPORT
if args.odt:
    template_file = args.odt
if args.csv:
    csv_file = args.csv

try:
    inserter \
        = InsertionProcess.InsertionProcess(template_file,use_loport, \
                                            symbolsdir=use_iconsdir, \
                                            pdfoutdir=use_pdfexport )
except:
    inserter.cleanup()
    logging.error("Can't connect LibreOffice, or open Document {}".format(template_file))
    exit(1)

atexit.register(inserter.cleanup)
logging.info("Use template {}, load icons from {} and export PDF to {}" \
             .format(template_file,use_iconsdir,use_pdfexport))

# pandas を使って CSV を読み込む。
# read_table はタブ切りを読み込み、dtype=str と指定することで
# 全てのレコードを文字列として読み込む
csvrecords = pd.read_table(csv_file,dtype=str)
# 全カラムが空値の行を削除する
csvrecords = csvrecords.dropna(how='all')
# NaN を空文字で置き換えておく
csvrecords = csvrecords.fillna('')
logging.info("Load csv{} done".format(csv_file))

try:
    inserter.main_loop_pandas(csvrecords)

finally:
    pass

logging.info("Done.")
