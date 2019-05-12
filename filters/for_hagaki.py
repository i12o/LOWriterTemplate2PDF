import os
from logging import getLogger

logging = getLogger(os.path.basename(__file__)).getChild(__name__)

def zip(num, data, obj):
    '''split Zip1, Zip2 and make Zip11,Zip12,Zip13... and Zip21,Zip22...'''
    for z in ['Zip1','Zip2','SZip1','SZip2']:
        if data.get(z):
            for i,s in enumerate(list(data.get(z))):
                data[z+str(i+1)] = s
    return True

# 一文字しかない名前を両端揃えするとなんかバランス悪い気がするので、調
# 正をデータの方で加えてみる
def one_char_name(num, data, obj):
    '''When (Family|First)Name[1-3] contains only 1 letter, this adds
    whitespace before and after that letter.'''

    for n in ['FamilyName1','FamilyName2','FamilyName3',
              'FirstName1','FirstName2','FirstName3',
    ]:
        if data.get(n):
            if len(data.get(n)) == 1:
                data[n] = ' ' + data.get(n) + ' '
    return True

def select_one_style(num,data,obj):

    pc_colname = '__meta.printcontrol'
    if data.get(pc_colname):
        return True

    logging.warn("Select page")
    # 宛先指定が一つの場合、テンプレートの1ページ目を使う。
    onlyonereceipt = True
    for n in ['FamilyName2','FirstName2',
              'FamilyName3','FirstName3' ]:
        if data.get(n) != "":
            onlyonereceipt = False
    if onlyonereceipt:
        data[pc_colname] = ['1']
        logging.debug("Page: {}".format(data[pc_colname]))
        return True

    samefamily = True
    for n in [ 'FamilyName2','FamilyName3' ]:
        logging.debug("samefamily: {} {}".format(n,data.get(n)))
        if data.get(n) != "":
            logging.debug("hit samefamily: {} {}".format(n,data.get(n)))
            samefamily = False
    if samefamily:
        data[pc_colname] = ['2']
        logging.debug("Page: {}".format(data[pc_colname]))
        return True

    # フォールバックデフォルト
    data[pc_colname] = ['3']
    logging.debug("Page: {}".format(data[pc_colname]))
    return True
