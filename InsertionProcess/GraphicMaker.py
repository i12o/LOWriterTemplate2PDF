# Use python-barcode

import sys
import os
import glob
from logging import getLogger
import base64
import io

import barcode
import pyqrcode

logging = getLogger(os.path.basename(__file__)).getChild(__name__)

# 'Name of algorithm to create image' : 'Code reference'
known_algorithm = {}

supported = {
    "EAN8"    : "EAN8",
    "EAN13"   : "EAN13",
    "EAN14"   : "EAN14",
    "CODE39"  : "Code39",
    "CODE128" : "Code128",
    "JAN"     : "JAN",
    "UPCA"    : "UPCA",
    "ISBN13"  : "ISBN13",
    "ISBN10"  : "ISBN10",
    "ISSN"    : "ISSN",
    "PZN"     : "PZN",
    "ITF"     : "ITF",
}

# If you need NW7/Codabar, there're barcode fonts for them.
# Just install these, and just change font of textfield to
# Codabar font.

def _barcode_binder(sig):
    def closure(x):
        return _barcode_generator(sig,x)
    return closure

for k,v in supported.items():
    known_algorithm[k] = _barcode_binder(v)

def _barcode_generator(barcode_type,value):
    bc = barcode.get(barcode_type,value)
    # As I can't see how to get binary expression of barcode image,
    # I use BytesIO.  Arenn't there more better way?
    fp = io.BytesIO()
    bc.write(fp)
    fp.seek(0)
    barcode_bin = fp.read()
    fp.close()
    return barcode_bin

def create_qrcode_svg(value):
    qrcode = pyqrcode.create(value)
    fp = io.BytesIO()
    qrcode.svg(fp,scale=4)
    fp.seek(0)
    image = fp.read()
    fp.close()
    return image

known_algorithm["QRCode"] = create_qrcode_svg

def get_graphic_maker_func(algoname):
    return known_algorithm.get(algoname)

def raw_png(filename):
    '''生のPNGファイルを読み込む。画像生成アルゴリズムではない'''
    with open(filename,mode='rb') as f:
        imgdata = f.read()
        f.close()
    if not imgdata:
        logging.warn("Can't read PNG file %s" %(filename,))
        return
    return imgdata

def make_list_of_pict_selection(path):
    '''pathで指定されたディレクトリのpng一覧から、リストを自動生成する'''
    '''ファイル名は、識別子.選択値.png という形式でなければならない'''
    logging.debug("Pickup icons from %s" %(path, ))
    tmplist = {}
    imagelist = {}
    for f in glob.glob(os.path.join(path,'*.[Pp][Nn][Gg]')):
        filename = os.path.basename(f)
        (basename,ext) = os.path.splitext(filename)
        dotsepped = basename.split(".")
        if len(dotsepped[1:]) == 0:
            logging.warning("%s contains png file which has invalid filename format %s" \
                            %(path, filename))
            logging.warning("Filename must be DESIG.VALUE.png")
            continue
        selval = '.'.join(dotsepped[1:])
        selkey = dotsepped[0]
        if tmplist.get(selkey):
            tmplist[selkey].append({ "filename": f, "val": selval})
        else:
            tmplist[selkey] = [ { "filename": f, "val": selval} ]
    for k in tmplist.keys():
        sortedents = sorted(tmplist[k],key=lambda e:e["val"])
        if [ i for i in sortedents if not str(i['val']).isdigit() ]:
            # 整数でない選択値が含まれている場合
            imagelist[k] = { j["val"]:j["filename"] for j in sortedents }
            imagelist[k]["default"] = sortedents[0]["val"]
        else:
            # 整数のみから構成されている場合
            # 整数選択値のみから構成される場合、リストにしようかと思ったんだが
            # 選択値として 0〜 ではなく、1〜とか、飛び番で指定される可能性を
            # 考えると、そのまま辞書の方が良い気もしてきたぞ。
            imagelist[k] = { j["val"]:j["filename"] for j in sortedents }
            imagelist[k]["default"] = sortedents[0]["val"]
    return imagelist
