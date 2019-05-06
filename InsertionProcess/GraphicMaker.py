# 画像イメージの生成処理を纏める
# UNO には依存しない

import sys
import os
import glob
import reportlab.graphics.barcode
from reportlab.lib.units import mm
import base64
import pyqrcode
import logging

# Name of algorithm to create image : Code reference
known_algorithm = {}

barcode_width = 640
barcode_height = 240


def create_ean13_png_image(value):
    return reportlab.graphics.barcode.createBarcodeImageInMemory(
        'EAN13', value = value, format='png',
        width = barcode_width, height = barcode_height, humanReadable =1)

known_algorithm["EAN13"] = create_ean13_png_image

# Generated image might have quality problem, so
# use of NW7/Codabar Generator is not recommended.
# It might be better using Barcode font for these.
def create_nw7_png_image(value):
    # ストップがあるため2文字増える
    l = len(value)+2
    # 3文字で14単位増え、mod 3 に従い 5,4,5 単位増える
    # サイレントエリアが5単位*2
    barunits = 14 * ( l // 3 ) + (0,5,9)[l % 3] + 10
    return reportlab.graphics.barcode.createBarcodeImageInMemory(
        'Codabar', value = value, format='png',
        width= barunits*16, height = barcode_height * 2, humanReadable =1)
#        barWidth=0.3*mm, height = barcode_height, humanReadable =1)

known_algorithm["NW7"] = create_nw7_png_image
known_algorithm["Codabar"] = create_nw7_png_image

def create_code39_png_image(value):
    return reportlab.graphics.barcode.createBarcodeImageInMemory(
        'Standard39', value = value, format='png',
        width = barcode_width, height = barcode_height, humanReadable =1)

known_algorithm["CODE39"] = create_code39_png_image

def create_extended39_png_image(value):
    return reportlab.graphics.barcode.createBarcodeImageInMemory(
        'Extended39', value = value, format='png',
        width = barcode_width, height = barcode_height, humanReadable =1)

known_algorithm["EXTENDED39"] = create_extended39_png_image

def create_code128_png_image(value):
    return reportlab.graphics.barcode.createBarcodeImageInMemory(
        'Code128', value = value, format='png',
        width = barcode_width * 2, height = barcode_height * 2, humanReadable =1)

known_algorithm["CODE128"] = create_code128_png_image

def create_qrcode_png_image(value):
    qrcode = pyqrcode.create(value)
    pngimage = base64.b64decode(qrcode.png_as_base64_str(scale=4))
    return pngimage

known_algorithm["QRCode"] = create_qrcode_png_image

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
