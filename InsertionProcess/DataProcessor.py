# 差し込みデータを受けて、どのような処理を行なうかのリストを作成する
# 言わばコンパイラ?
# LibreOffice に依存しない処理

import os
import re
from logging import getLogger

nameseparator = '.'

logger = getLogger(os.path.basename(__file__)).getChild(__name__)

# 辞書データ、差し替え画像の名前一覧、帳票差し込みプロセッサモジュールを
# 受けて処理リストを作成。
def create_list_of_processing(record, list_of_imagename, processor):
    return _compile_data_to_processlist(
        record,
        list_of_imagename
    )

# 辞書データを受けて、その一つ一つから行なうべき処理リストを作成
# 存在する画像のリストなど必要。
# key of dictionary must be one of:
#   * uservariable name
#     ex. Address1, Name1, Name2 ...
#   * Image replacing designator
#     * LABEL.Toggle
#       Image named with LABEL.ICON, is replaced with ICON.value
#         ex. Fragile.Toggle, Logo.Toggle ...
#     * LABEL.Selection
#       Series of image named with LABEL.value.ICON is replaced
#       depending of value of field.
#         ex. ShipMethod.Selection, Payment.Selection ...
#     * LABEL.ImageGenerator
#       Image named LABEL replaced with Generated image,
#       specified ImageGenarator.
#         ex. ShipNo.CODE128, QueryURL.QRCode ...

# imagenames contains all Labels of GraphicObject, in document.
# algolist is list of graphic generator implemented in GraphicMaker.py
def _compile_data_to_processlist(data, imagenames):
    proclist = []
    for (key,value) in data.items():
        labelelems = key.split(nameseparator)
        #logger.debug(labelelems)
        if (len(labelelems) == 1):
            proclist.append({
                "Type": 'UserVar',
                'Name': key,
                'Value': value,
            })
        elif (len(labelelems) >= 2):
            if (labelelems[1] == 'Toggle'):
                r = _toggle_reform(key,labelelems[0],value,imagenames)
                if r:
                    proclist.extend(r)
            elif (labelelems[1] == 'Selection'):
                r = _select_reform(key,labelelems[0],value,imagenames)
                if r:
                    proclist.extend(r)
            else:
                r = _gen_images(key,labelelems[1],value,imagenames)
                if r:
                    proclist.extend(r)
    return(proclist)

def _gen_images(label, algo, value, images):
    """Generate image, by known as alogo"""
    if [ e for e in images if e == label]:
        return [{
            'Type': 'GenerateImage',
            'ImageName': label,
            'Algorithm': algo,
            'Value': value
        }]

def _toggle_reform(label,paraname,value,images):
    """For Toggle button type(Not limited to on/off). Parameter Name is Param.Toggle"""
    r = []
    for imagename in [e for e in images if e.startswith(paraname)]:
        nameelem = imagename.split(nameseparator)
        imglist = nameelem[1] if len(nameelem)>1 else None
        r.append({
            'Type': 'ImageSelect',
            'ImageName': imagename,
            'Value': value,
            'ImageList': imglist
        })
    return r

def _select_reform(label,paraname,value,images):
    """One or more items in series of selection are toggled as on. For key Param.Selection"""
    r = []
    if not (isinstance(value,list)) and not (isinstance(value,tuple)):
        value = re.split('[, ]',value)
    for imagename in [e for e in images if e.startswith(paraname)]:
        nameelem = imagename.split(nameseparator)
        if len(nameelem) < 2:
            # 画像名に選択肢の値が提供されていない
            logger.warn("%s exists for selection %s, but no value specified" \
                          % (imagename,paraname))
            continue
        selectval = nameelem[1]
        imglistname = nameelem[2] if len(nameelem)>2 else None
        thisval = 0;
        if selectval in value:
            thisval = 1
        r.append({
            'Type': 'ImageSelect',
            'ImageName': imagename,
            'Value': thisval,
            'ImageList': imglistname
        })
    return r

if __name__ == "__main__":
    logger = getLogger(__file__)
    testdata = {
        'A':1,
        'B':'B',
        'C.Toggle': 'Ca',
        'C.X': 'Cb',
        'D.Selection': 'D',
        'D2.Selection': ['D','A'],
        }
    imagenames = [
        'C.la',
        'C.X',
        'B.qr',
        'D.A.Check',
        'D.B.Check',
        'D.C.Check',
        'D.D.Check',
        'D2.A.Check',
        'D2.B.Check',
        'D2.C.Check',
        'D2.D.Check',
    ]
    r = preprocess_data(testdata,imagenames,['X'])
    print()
    for i in r:
        print(i)
