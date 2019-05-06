# 帳票テンプレートへの差し込み用の処理
# 実際に処理を行うため、uno(ドキュメントオブジェクト)に依存する。
# uno の前処理などは行なわない。

import sys
import os
import warnings
import logging
import InsertionProcess.GraphicMaker as GraphicMaker

# PIL は SVG に対応していないので、PNG 画像を使う
# icon_list = {
#     "Check": [
#         "checkbox.png",
#         "checkbox_checked.png"
#     ],
#     "ValueSel": {               # 値に対応する画像がある
#         "0": "checkbox.png",
#         "1": "checkbox_checked.png",
#         "default": "checkbox.png", # default を必ず定義。
#     }
# }


known_algorithm = GraphicMaker.known_algorithm

# ファイル名を礎としてグラフィックオブジェクトをキャッシュする用

# path 以下にある、画像リストを読み込む。


class Processor:
    def __init__(self,_document,icons_dir):
        self.document = _document # LOif.Document
        self.graphic_object_cache = {}
        self.setup_icon_list(icons_dir)
        logging.info("Initialized")

    def setup_icon_list(self,path):
        self.icon_list = GraphicMaker.make_list_of_pict_selection(path)

    def _pick_filename_from_icon_list(self,listname,value):
        entity = self.icon_list.get(listname) \
                 or self.icon_list.get('Check')
        path = None
        v = str(value)

        logging.debug("Entity type: {} v: {} value {}".format(type(entity),v,value))
        if isinstance(entity,list):
            v = int(v) if str.isdecimal(v) else 0
            path = entity[v] if len(entity) > v else entity[0]
        else:
            # 辞書しかないと想定しているが。
            path = entity.get(v) or entity.get("default")
        return path

    def get_graphic_object_from_icons(self,listname,value):
        '''指定リストの指定値に該当するグラフィックオブジェクト(iconsのやつ)を取得する'''
        logging.debug("Given %s,%s" %(listname,value))
        filename = self._pick_filename_from_icon_list(listname,value)
        logging.debug("Filename %s" %(filename,))
        gobj = self.graphic_object_cache.get(filename)
        if gobj:
            logging.debug("Got from cache")
            return gobj
        imgdata = GraphicMaker.raw_png(filename)
        if not imgdata:
            logging.error("Image file %s not found for %s, %s" \
                          %(filename, listname, value))
            return
        logging.debug("Loaded from file %s" %(filename,))
        gobj = self.document.create_graphic_from_binary(imgdata)
        self.graphic_object_cache[filename] = gobj
        return gobj


    def insertion_process(self,proclist):
        '''uno関連のオブジェクトと、処理記述リストを受け取って、帳票への差し込み処理を行なう'''
        # 処理速度を速くするために、画面更新を止める
        self.document.lock_controller()
        for item in proclist:
            if item["Type"] == "UserVar":
                self.document.build_user_variable(
                    item["Name"],
                    str(item["Value"]),
                    nocreate = 1
                )
            elif item["Type"] == "GenerateImage":
                func = known_algorithm[item["Algorithm"]]
                if func:
                    g = func(item["Value"])
                    newobj = self.document.create_graphic_from_binary(g)
                    if newobj:
                        self.document.replace_graphic(
                            item["ImageName"],
                            newobj,
                            NameDic=self.document.grouped_graphics
                        )
                    else:
                        logging.warning("Failed to create Image %s.%s with %s" \
                                        %(item["Name"],item["Algorithm"],item["Value"]))
                else:
                    logging.warning("No such Image Generator known %s, for %s" \
                                    %(item["Algorithm"],item["Name"]))
            elif item["Type"] == "ImageSelect":
                gobj = self.get_graphic_object_from_icons(item["ImageList"], item["Value"])
                self.document.replace_graphic(
                    item["ImageName"],
                    gobj,
                    NameDic=self.document.grouped_graphics
                )
            else:
                logging.warning("Unsupported processing type: %s" %(item["Type"],))
        self.document.unlock_controller()
