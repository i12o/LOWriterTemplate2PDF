import os
import sys
import csv
import time
from datetime import datetime
import re
import string
from logging import getLogger
import atexit
from collections.abc import Iterable

import uno
import unohelper
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
# filenamer(IndexNumberOfData, DicOfData, self)
#   PDFとして出力するファイル名を返す関数

# give '.odt' as template_file,
# LO connection desc as loportstr
# If input data convertion required, give converter function list as data_converter.
# This function must destructively modify passed dictionary itself, and return True.
# When it returns False, that data is skipped and no PDF exported.
# If you change Exported PDF's filename instead of first field value of input data,
# pass function which returns filename from single tuple of input data as filenamer.
#
# spillfix:
#    None, False, 0 => Never try
#    2 => Try spillfix, and it fails to fix, no PDF is written
#    True, not 2 => Try spillfix, and output PDF anyway.
class InsertionProcess:
    def __init__(self,
                 template_file, # filename of odt
                 loportstr,     # --accept parameter of LO
                 *,
                 symbolsdir=SYMBOLSDIRPATH, # Relative to template
                 data_converter=None,       # Data converter
                 filenamer=None,            # Filename creater
                 pdfoutdir=PDFEXPORTDIR,    # PDF output dir
                 spillfix=None, # Experimental: do spillfix
                 printer=None,  # Directly print to Specified Printer
    ):
        self._data_converter = None
        if data_converter:
            if isinstance(data_converter,Iterable):
                self._data_converter = data_converter
            else:
                self._data_converter = (data_converter,)
        self._filenamer = filenamer
        self._pdfdir = pdfoutdir
        self._do_cleanup = False
        self._spillfix = spillfix
        self._printer = printer
        self._direct_print = False

        logging.debug("Connecting libreoffice with {}".format(loportstr))
        self.context = LOif.ConnectLO(loportstr)
        if not self.context:
            logging.error("Failed to LibreOffice service, %s" %(loportstr,))
            exit(1)

        self.document = LOif.Document(self.context,filename=template_file)

        if template_file:
            self._do_cleanup = True
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

        if self._printer is not None:
            logging.info("Printout directly to printer named as {}".format(self._printer))
            self._direct_print = True
            self.set_printer()

    # プリンターを指定する。空文字が指定されていればそのままいじらない。
    # 文字列が指定されていた場合、プリンターを指定のプリンタにする。
    def set_printer(self):
        if self._printer:
            # Usable property:
            #   Name: Name of printer in string
            #   PaperOrientation: CSS.view.PaperOrientation.PORTRAIT
            #     or  CSS.view.PaperOrientation.LANDSCAPE. enum
            #   PaperFormat: CSS.view.PaperFormat.A4、etc. enum
            #   PaperSize: struct CSS.awt.Size.  size.Width, size.Height
            #     must be int, size in mm x 100 times.
            props = { 'Name': self._printer }
            self.document.set_printer_property(props)

    def insert_record_process(self,record):
        self.document.reset(spillfix = self._spillfix)
        graphicgrpnames = self.graphic_variant_dic.keys() \
            if self.graphic_variant_dic else ()
        proclist = PreProcess.create_list_of_processing(
            record, graphicgrpnames, DoInsertion)
        self.processor.insertion_process(proclist)

    def insert_and_export(self,record,num):
        logging.debug("Insert {} : record: {}".format(num,record))
        self.insert_record_process(record)
        if self._spillfix:
            spillresult = self.document.do_spillfix()
            logging.debug("spillresult {}".format(spillresult))
            if spillresult is None:
                logging.debug('Tried spillfix, but nothing changed')
            elif spillresult is False:
                logging.warning('Tried spillfix, but some textframes still spill out')
                if self._spillfix == 2:
                    logging.error('record number {} has unfixable spillout frame. Skipped' \
                                  .format(num+1))
                    return
            else:
                logging.debug('Successfully spillfixed {}'.format(num))

        # PageRangeとして '1-3; 1;' が与えられたら、テンプレートの
        # 1,2,3,1 ページからなる4ページのPDFが生成される。
        _fdata = {}
        if record.get('__meta.printcontrol'):
            if self._direct_print:
                _fdata = { 'Pages': record['__meta.printcontrol'] }
            else:
                _fdata = { 'PageRange': record['__meta.printcontrol'] }
        if self._filenamer:
            pdffilename = self._filenamer(num,record,self)
        else:
            pdffilename = '{:04d}.pdf'.format(num+1)
        if self._direct_print:
            self.document.send_to_printer(
                filterdata = _fdata,
            )
        else:
            self.document.export_as_pdf(
                os.path.join(self._pdfdir,pdffilename),
                filterdata = _fdata,
            )

    def transform_record(self, num, record):
        # transform single line of data, it may processed with optional
        # data_converter
        data = record.copy()
        logging.debug(data)
        self.auto_complement_algorithm_field(data)
        logging.debug("After auto_complement {}".format(data))
        if self._data_converter:
            for conv in self._data_converter:
                logging.debug('Data Conversion Process {}'.format(conv.__name__))
                if not conv(num,data,self):
                    return
        return data

    # データレコードに ShippingNo だけがあり、テンプレート文書の画像
    # ラベルとして ShippingNo.EAN13 と ShippingNo.EXTENDED39 がある
    # 場合に、ShippingNo.EAN13、ShiipingNo.EXTENDED39 のレコードをデータ
    # レコードに生成する。
    def auto_complement_algorithm_field(self, data):
        '''データレコードと画像名の辞書から、生成すべき画像種のデータレコードを自動生成する'''
        graphic_algos = list(GraphicMaker.known_algorithm.keys()) + ["URL"]
        for dk in { s for s in data.keys() if len(s.split('.',1)) == 1 }:
            for alg in graphic_algos:
                newdk = '%s.%s' %(dk,alg)
                if data.get(newdk) \
                   or not (self.graphic_variant_dic \
                           and self.graphic_variant_dic.get(newdk)):
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
