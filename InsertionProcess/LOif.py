# UNO を使った一般的な処理をまとめる

import sys
import os
import uno
import unohelper
from com.sun.star.beans import PropertyValue
import time
from logging import getLogger

logging = getLogger(os.path.basename(__file__)).getChild(__name__)

CSS = 'com.sun.star'

ooport = "socket,host=localhost,port=2083;urp;"

def resolverarg(portstr=ooport):
    r = "uno:" + ooport + "StarOffice.ComponentContext"
    logging.debug("resolverarg: %s" %(r,))
    return r

def ConnectLO(portstr=ooport):
    local = uno.getComponentContext()
    resolver = local.ServiceManager \
                    .createInstanceWithContext(CSS + ".bridge.UnoUrlResolver", local)
    logging.debug("Resolver: {}".format(resolver))

    try:
        # 接続できるか?
        logging.debug('Connecting to LibreOffice')
        context = resolver.resolve(resolverarg(portstr))

    except Exception as e :
        # エラーが発生した場合
        if 'connect to socket' in str(e):
            # 起動する
            os.system('''soffice --accept="''' + portstr + '"')
            # LibreOfficeが起動するまで三秒待つ
            time.sleep(3)
            # 再度、接続できるか試す
            logging.debug('LibreOffice was not opened. I have openned it and try connecting')
            context = resolver.resolve(resolverarg())

    if not context:
        logging.critical("Can't connect to LibreOffice, with {}".format(portstr))
        logging.critical("Check your LibreOffice setting.")
        exit(1)

    return context

class Document:
    def __init__(self,
                 _context,       # uno context
                 **args ):
        # optional : 'filename', open filename or current active LO document is used.

        self.context = _context
        self.smgr = self.context.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext(CSS + ".frame.Desktop", self.context)
        logging.debug('Document args: {}'.format(args))
        if not args.get('filename'):
            self.document = self.desktop.CurrentComponent
            self.templatefile = None
            self.template_url = None
        else:
            self.templatefile = os.path.abspath(args['filename'])
            self.template_url = unohelper.systemPathToFileUrl(self.templatefile)
            self.document = \
                self.desktop.loadComponentFromURL(self.template_url ,"_blank", 0, ())

        if not self.document:
            logging.critical("'Can't open document filename={}".format(args.get("filename")))
            exit(1)

        self.drawpage = self.document.getDrawPage()
        self.form = self.drawpage.getForms()
        self.graphicprovider = self.smgr.createInstanceWithContext(
            CSS + ".graphic.GraphicProvider", self.context
        )
        (self.inital_graphics, self.grouped_graphics) = \
            self.get_all_graphics_in_doc()
        self.all_textfields = self.get_all_textfield_in_doc()
        return;

    # export document in PDF
    # If you want to export only pages 1,2 of document, passing filterdata
    # as this:
    #   exportas_pdf('/tmp/testout.pdf',filterdata={'PageRange':'1-2'})
    # will do the job.
    def export_as_pdf(self,filename,filterdata=None):
        exportprop = [
            PropertyValue("FilterName",0,"writer_pdf_Export",0),
        ]
        if filterdata:
            fd = []
            for (k,v) in filterdata.items():
                fd.append(PropertyValue(k,0,v,0))
                fdp = PropertyValue("FilterData",0,
                                    uno.Any("[]com.sun.star.beans.PropertyValue",
                                            tuple(fd)), 0)
                exportprop.append(fdp)
        exportpdfname = unohelper.systemPathToFileUrl(
            os.path.abspath( filename )
        )
        try:
            logging.debug("Export as {}".format(exportpdfname))
            self.document.storeToURL(exportpdfname,exportprop)
        except:
            logging.error("Can't export {}".format(exportpdfname))

    def reset(self):
        # TODO: reset modified parameters to original
        pass

    def cleanup(self):
        self.document.dispose()

    def lock_controller(self):
        self.i_locked_it = 0
        logging.debug("Lock controller is locked?{}".format(self.document.hasControllersLocked()))
        if not self.document.hasControllersLocked():
            self.document.lockControllers()
            self.i_locked_it = 1;

    def unlock_controller(self):
        if self.i_locked_it:
            self.document.unlockControllers()
            self.i_locked_it = 0

    def get_all_graphics_in_doc(self):
        '''ドキュメント中の全画像名に対するオブジェクトの辞書と、バリアントのリストを返す'''
        graphiclist = {}
        groupedlist = None
        gobjects = self.document.GraphicObjects
        if gobjects:
            for n in gobjects.ElementNames:
                graphiclist[n] = gobjects.getByName(n).Graphic
            groupedlist = self._make_namespec_to_real_graphic_names(graphiclist)
        return (graphiclist,groupedlist)

    def _make_namespec_to_real_graphic_names(self,graphiclist):
        '''If GraphicObject has Label, "Image,0", "Image,1"..., group them into'''
        '''"Image" group.'''
        groupedlist = {}
        variant_exists = 0;
        for k in graphiclist.keys():
            # 画像の名前を","を区切りとして最大2つに分割
            nl = k.split(",",1)
            if groupedlist.get(nl[0]):
                groupedlist[nl[0]].append(k)
            else:
                groupedlist[nl[0]] = [k,]
            if len(nl)>1:       # ,サフィックスが存在する場合
                variant_exists = 1
        return groupedlist

    def create_graphic_from_binary(self,binary):
        """グラフィックオブジェクトを生成"""
        istream = self.smgr.createInstanceWithContext(
            CSS + ".io.SequenceInputStream",self.context
        )
        istream.initialize((uno.ByteSequence(binary),))
        logging.debug("Istream created %s" %(istream))
        gobj = self.graphicprovider.queryGraphic((
            PropertyValue("InputStream",0,istream,0),
        ))
        logging.debug("queryGraphic %s" %(gobj,))
        return gobj

    # pass first returned value of get_all_graphics_in_doc()
    def reset_all_graphic_in_doc(self):
        gobjects = self.document.GraphicObjects
        for (name,gobj) in self.initial_graphics.items():
            gobjects.getByName(name).Graphic = gobj

    def get_all_textfield_in_doc(self):
        '''ユーザ定義変数によるテキストフィールドを列挙する'''
        fieldnamelist = {}
        fields = self.document.TextFields
        if fields:
            for i in fields.createEnumeration():
                if i.supportsService(CSS + '.text.TextField.User'):
                    fieldnamelist[i.TextFieldMaster.Name] = i.TextFieldMaster.Content
        return fieldnamelist

    def build_user_variable(self, name, value, nocreate=0):
        prefix = CSS + ".text.FieldMaster.User"
        master = self.document.TextFieldMasters
        vname = '%s.%s' % (prefix,name)
        if master.hasByName(vname):
            # variable already defined
            var = master.getByName(vname)
            var.Content = value
        else:
            if nocreate:
                logging.warning("UserVariable %s not exist" %(name,))
                return
            var = self.document.createInstance(CSS + '.text.FieldMaster.User')
            var.Name = name
            var.Content = value
        return var

    # This replaces GraphicObject.Graphic Labeled as imagename, with newobj
    # If group of image list given, it replaces all members of 'imagename'.
    # group of image is given as NameDic, it's got by get_all_graphics_in_doc,
    # 2nd return value.
    def replace_graphic(self,imagename,newobj,NameDic=None):
        nlist = [imagename]
        if NameDic and NameDic.get(imagename):
            nlist = NameDic[imagename]
        for n in nlist:
            oldgraphic = self.document.GraphicObjects.getByName(n)
            if oldgraphic:
                oldgraphic.Graphic = newobj
            else:
                logging.warning("Failed to find Image %s" %(n,))

    # 指定のテキストフレームが固定サイズである場合に、中身がはみ出しているかどうか、
    # 自動サイズ変更に切り替えてみて大きさが変化するかどうかで調べる。
    # よって、このテキストフレームの内側にある要素ではみ出しが生じていても
    # それを検知はできない。
    def check_textframe_content_spilled_out(self,textframe):
        is_overflow = 0
        if textframe.FrameIsAutomaticHeight or textframe.WidthType != 1:
            # Automatic Size.  It'll not spill out.
            is_overflow = 2
            return False
        prevlayout = textframe.LayoutSize
        prevwidthtype = textframe.WidthType
        prevautoheight = textframe.FrameIsAutomaticHeight
        textframe.FrameIsAutomaticHeight = True
        if prevlayout.Height < textframe.LayoutSize.Height:
            is_overflow = 1
            textframe.FrameIsAutomaticHeight = prevautoheight
            textframe.WidthType = prevwidthtype
        if is_overflow == 1:
            return True

    # テキストフレーム内部にあるテキストのフォントサイズを調整する。
    # テキストフレーム以外のパラメータを指定せずに呼び出すと、その
    # テキストフレーム内部のテキスト要素の現在のフォントサイズが返される
    # TODO: Before applying this, text size must be saved, and
    #   resettable.
    #   When this textframe contains other than Text, it'll raise
    #   exception not having CharHeight property
    def textframe_text_setsize(self,textframe,size=None,reset=None):
        original = []
        prop_heights = ('CharHeight', 'CharHeightAsian','CharHeightComplex')
        index = 0
        for t in textframe.Text.createEnumeration():
            prev = {}
            for ch in prop_heights:
                if reset:
                    t.setPropertyValue(ch,reset[index][ch])
                else:
                    prev[ch] = t.getPropertyValue(ch)
                    if size:
                        t.setPropertyValue(ch,size)
            if prev.keys():
                original.append(prev)
        if len(original):
            return original
