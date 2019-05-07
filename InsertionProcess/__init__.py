import os
import sys
import csv
import time
from datetime import datetime
import re
import string
from logging import getLogger
import atexit

import uno
import unohelper
from com.sun.star.beans import PropertyValue
import pandas as pd

import InsertionProcess.DoInsertion
import InsertionProcess.PreProcess
import InsertionProcess.LOif
import InsertionProcess.GraphicMaker

logging = getLogger(os.path.basename(__file__)).getChild(__name__)

SYMBOLSDIRPATH = 'icons'
PDFEXPORTDIR = 'pdfout'

# data_converter(IndexNumberOfData,DicOfData,self)
#   破壊的に DicOfData を変更し、データに対する操作を加える関数を指定。
#   True を返すと差し込みを行ない、False ならこのデータに対する処理はスキップする
# filenamer(IndexNumberOfData, DicOfData, self, [pcr=PageRange, pcindex=INDEX])
#   PDFとして出力するファイル名を返す関数
#   差し込みデータの、__meta.printcontorl というレコードにリスト、タプルなどを渡し
#   た場合、その要素を PageRange として PDF出力を行なう。
#   例えば [ '1','2','1-2' ] という値があった場合
#     1ページのみ、2ページのみ、1-2ページ
#   という3つのPDFが作成される。この一つ一つのPageRange指定と、インデックスが
#   pcrとpcindexとして渡される。
#   これらを使ってファイル名を生成するようにしておかないと、折角印刷ページ指示を
#   しても同じファイルに書き出される事になってしまう。

# give '.odt' as template_file,
# LO connection desc as loportstr
# If input data convertion required, give converter function as data_converter.
# This function must destructively modify passed dictionary itself.
# If you change Exported PDF's filename instead of first field value of input data,
# pass function which returns filename from single tuple of input data as filenamer.
class InsertionProcess:
    def __init__(self,
                 template_file, # filename of odt
                 loportstr,     # --accept parameter of LO
                 symbolsdir=SYMBOLSDIRPATH, # Relative to template
                 data_converter=None,       # Data converter
                 filenamer=None,            # Filename creater
                 pdfoutdir=PDFEXPORTDIR ):  # PDF output dir
        self._data_converter = data_converter
        self._filenamer = filenamer
        self._pdfdir = pdfoutdir

        logging.debug("Connecting libreoffice with {}".format(loportstr))
        self.context = LOif.ConnectLO(loportstr)
        if not self.context:
            logging.error("Failed to LibreOffice service, %s" %(loportstr,))
            exit(1)

        self.document = LOif.Document(self.context,filename=template_file)

        if template_file:
            self._do_cleanup = 1
            self.template_base = os.path.dirname(template_file)
            logging.debug("Opened {} as templte file".format(template_file))
        else:
            self.template_base = './'
            logging.debug("Use current active document as template")

        self._symbolsdir = os.path.join(self.template_base, symbolsdir)
        logging.debug("icons will be read from {} directory".format(self._symbolsdir))

        (self.initial_graphics,self.graphic_variant_dic) \
            = self.document.get_all_graphics_in_doc()
        self.processor = DoInsertion.Processor(self.document,self._symbolsdir)

    def insert_record_process(self,record):
        self.document.reset()
        proclist = PreProcess.create_list_of_processing(
            record, self.graphic_variant_dic.keys(), DoInsertion)
        self.processor.insertion_process(proclist)

    def insert_and_export(self,record,num):
        logging.debug("Insert {} : record: {}".format(num,record))
        self.insert_record_process(record)

        # PageRangeのリスト。
        #   [ '1-2' ,'1' ]
        # が与えられたら、1-2ページの出力、1ページの出力が生成される。
        if record.get('__meta.printcontrol'):
            pcindex = 1;
            for c in record['__meta.printcontrol']:
                _fdata = {'PageRange':c}
                if self._filenamer:
                    pdffilename = self._filenamer(num,record,self,
                                                  pcr=c, pcindex=pcindex)
                else:
                    pdffilename = '{:04d}_{:02d}.pdf'.format(num+1,pcindex)
                self.document.export_as_pdf(
                    os.path.join(self._pdfdir,pdffilename),
                    filterdata = _fdata
                )
                pcindex+=1
        else:
            if self._filenamer:
                pdffilename = self._filenamer(num,record,self)
            else:
                pdffilename = '{:04d}.pdf'.format(num+1)
            self.document.export_as_pdf(
                os.path.join(self._pdfdir,pdffilename)
            )

    def transform_record(self, num, record):
        # transform single line of data, it may processed with optional
        # data_converter
        data = record.copy()
        logging.debug(data)
        self.auto_complement_algorithm_field(data)
        logging.debug("After auto_complement {}".format(data))
        if self._data_converter:
            if self._data_converter(num,data,self):
                return data
            else:
                return
        return data

    # データレコードに ShippingNo だけがあり、テンプレート文書の画像
    # ラベルとして ShippingNo.EAN13 と ShippingNo.EXTENDED39 がある
    # 場合に、ShippingNo.EAN13、ShiipingNo.EXTENDED39 のレコードをデータ
    # レコードに生成する。
    def auto_complement_algorithm_field(self, data):
        '''データレコードと画像名の辞書から、生成すべき画像種のデータレコードを自動生成する'''
        for dk in { s for s in data.keys() if len(s.split('.',1)) == 1 }:
            for alg in GraphicMaker.known_algorithm:
                newdk = '%s.%s' %(dk,alg)
                if data.get(newdk) or not self.graphic_variant_dic.get(newdk):
                    continue
                data[newdk] = data[dk]
                logging.debug("Complemented %s as %s" %(newdk,data[newdk]))

    def main_loop_pandas(self,datarecords):
        #Lets start the loop
        logging.debug('Starting to Loop Through Panda Records')
        # 処理速度を速くするために、画面更新を止める
        self.document.lock_controller()
        for num,row in datarecords.iterrows():
            record = dict(row)
            procrecord = self.transform_record(num,record)
            if procrecord:
                self.insert_and_export(procrecord, num)
        self.document.unlock_controller()

    def main_loop_list_of_dict(self,datarecords):
        #Lets start the loop
        logging.debug('Starting to Loop Through list of dict')
        # 処理速度を速くするために、画面更新を止める
        self.document.lock_controller()
        num = 0
        for record in datarecords:
            procrecord = self.transform_record(num,record)
            if procrecord:
                self.insert_and_export(procrecord, num)
            num += 1
        self.document.unlock_controller()

    def cleanup(self):
        if self._do_cleanup:
            self.document.cleanup()
