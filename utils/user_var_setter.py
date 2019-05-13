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
import InsertionProcess.common_options

# このディレクトリ下にチェックボックスなどの画像があるとする。
# テンプレートファイルからの相対でも、絶対パスでも良い
SYMBOLSDIRPATH = 'icons'
PDFEXPORTDIR = "./pdfout"
LOPORT="socket,host=localhost,port=2083;urp;"

optionvalues = {
    'loglevel' : logging.WARNING,
    'use_iconsdir' : SYMBOLSDIRPATH,
    'use_loport' : LOPORT,
    'use_pdfexport' : PDFEXPORTDIR,
    'template_file' :  None,
    'csv_file' : None,
    'spillfix' : None,
    'filenamer': None,
    'converters' : None,
}

parser = InsertionProcess.common_options.args_setup(
    "Read CSV and set uservariables to Template",
    optionvalues
)
parser.add_argument('-i','--index',
                    help="Set N-th of data, 0 origin",
                    type=int, default=0)

args = InsertionProcess.common_options.args_parse(parser,optionvalues)

index = args.index

logging.basicConfig(level = optionvalues['loglevel'])


try:
    inserter \
        = InsertionProcess.InsertionProcess(optionvalues['template_file'],
                                            optionvalues['use_loport'],
                                            symbolsdir=optionvalues['use_iconsdir'],
                                            pdfoutdir=optionvalues['use_pdfexport'],
                                            spillfix=optionvalues['spillfix'],
                                            filenamer=optionvalues['filenamer'],
                                            data_converter=optionvalues['converters'],
        )

except:
    inserter.cleanup()
    logging.error("Can't connect LibreOffice, or open Document {}" \
                  .format(optionvalues['template_file']))
    exit(1)

atexit.register(inserter.cleanup)

csvrecords = pd.read_table(optionvalues['csv_file'],dtype=str)
# 全カラムが空値の行を削除する
csvrecords = csvrecords.dropna(how='all')
# NaN を空文字で置き換えておく
csvrecords = csvrecords.fillna('')
logging.info("Load csv{} done".format(optionvalues['csv_file']))

d = dict(csvrecords.iloc[index])
for k,v in d.items():
    k = re.sub("\..*$","",k)
    inserter.document.build_user_variable(k,v)
    print("Set {} as {}".format(k,v))

logging.info("Done.")
