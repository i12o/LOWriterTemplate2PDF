# UNO を使った一般的な処理をまとめる

import sys
import os
import time
import re
from logging import getLogger

import uno
import unohelper
from com.sun.star.beans import PropertyValue

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
        # Get spillfix required frame names, and directive options
        self.spillfix_required = self.find_spillfix_required_textframes()
        # store each TextFrame's CharHeight* initial status
        self.initial_spillfix_frame_status = {}
        if self.spillfix_required:
            for name,frame in self.spillfix_required.items():
                (info,dummy) = self.textframe_text_setsize(frame['textframe'])
                self.initial_spillfix_frame_status[name] = info
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

    def reset(self,*,spillfix=False):
        if spillfix:
            # Reset modified parameters(TextSize) to original
            for name,dic in self.spillfix_required.items():
                self.textframe_text_setsize(dic['textframe'],
                                            Reset=self.initial_spillfix_frame_status[name])
                dic['textframe'].FrameIsAutomaticHeight = dic['FrameIsAutomaticHeight']
                dic['textframe'].WidthType = dic['WidthType']

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

    # Find TextFrames its description contains 'spillfix:something'
    # Optional something must be float.  Shrinked TextSize must be
    # over this limit.
    #   ex. spillfix:8 => Smallest TextSize is 8
    #       spillfix   => Smallest TextSize is 5(default)
    #                       specified in spillfix_textframe's LimitSize.
    def find_spillfix_required_textframes(self):
        spillfix = {}
        for tfn in self.document.TextFrames.ElementNames:
            tf = self.document.TextFrames.getByName(tfn)
            is_spillfix = re.search("spillfix(:([-a-zA-Z0-9_.]+))?",tf.Description)
            if is_spillfix:
                lim = None
                d = is_spillfix.group(2)
                try:
                    lim = float(d)
                except:
                    pass

                spillfix[tfn] = {
                    "directive": lim,
                    "textframe": tf,
                    # Resetting fontsize make FrameIsAutomaticHeight be True,
                    # so record previous status.
                    'FrameIsAutomaticHeight': tf.FrameIsAutomaticHeight,
                    'WidthType': tf.WidthType,
                }
        return spillfix

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

    def _get_textframe_contents(self,textframe):
        contents = []
        for t in textframe.Text.createEnumeration():
            contents.append(t)
        return contents

    # テキストフレーム内部にあるテキストのフォントサイズを調整する。
    # テキストフレーム以外のパラメータを指定せずに呼び出すと、その
    # テキストフレーム内部のテキスト要素の現在のフォントサイズが返される
    # とりあえず SizeDiff として負の float を指定すると、CharHeigt*
    # を全て +SizeDiff する。
    # フレーム内部に複数の Text オブジェクト が存在する場合、全ての
    # オブジェクト一律にサイズを変える。
    # TODO:
    #    フレームに対して Text.createEnumeration() して得られるのは、
    #   おそらく段落。段落に対してサイズ変更をすると、段落内の文字単位
    #   で指定しているサイズも上書きされ、このメソッドではそれらの情報
    #   を取れない。
    #    このため、Reset してもフレーム内のサイズ情報は失われる。
    def textframe_text_setsize(self,textframe,SizeDiff=None,Reset=None,AlertLimit=1.0):
        original = []
        prefix = 'CharHeight'
        styleprops = ('', 'Asian','Complex')
        index = 0
        if SizeDiff and SizeDiff >=0:
            logging.warn("SizeDiff must negative number, given {}".format(SizeDiff))
            return None,True
        under_alert = False
        for t in textframe.Text.createEnumeration():
            prev = {}
            for ch in styleprops:
                if t.supportsService(CSS + '.style.CharacterProperties' + ch):
                    if Reset:
                        t.setPropertyValue(prefix + ch,
                                           Reset[index][ch])
                    else:
                        prev[ch] = t.getPropertyValue(prefix + ch)
                        if SizeDiff:
                            newsize = prev[ch] + SizeDiff
                            if newsize <= AlertLimit or newsize <=0:
                                under_alert = True
                                newsize = AlertLimit if AlertLimit > 0 else 1
                            t.setPropertyValue(prefix + ch, newsize)
            if prev.keys():
                original.append(prev)
        if len(original):
            return original,under_alert

    def spillfix_textframe(self,textframe,LimitSize=5,DeltaStep=-0.5):
        '''This shrinks FontSize within textframe by DeltaStep, until
           contents of textframe doesn't spill out textframe.
           It returns None, when no shrinking done, False, when
           contents doesn't fit in textframe til LimitSize, and
           True when successfully fit content.'''
        status = None
        while self.check_textframe_content_spilled_out(textframe):
            (prevsize,reach_limit) = \
                self.textframe_text_setsize(textframe,
                                            SizeDiff=DeltaStep,
                                            AlertLimit=LimitSize)
            if reach_limit:
                if self.check_textframe_content_spilled_out(textframe):
                    status = False
                else:
                    status = True
                break
            status = True
        return status

    def do_spillfix(self,*,step=-0.5):
        '''This does spillfix for required textframes.  Nothing chagned,
           returns None.  All successfully spillfixed, returns True.
           Some of textframes are still spilling out, returns False.'''
        someframe_changed = False
        all_fit = True
        for name,d in self.spillfix_required.items():
            optarg = {'DeltaStep':step}
            if d['directive']:
                optarg['LimitSize'] = d['directive']
            result = self.spillfix_textframe(d['textframe'],**optarg)
            if not result is None:
                someframe_changed = True
                if result is False:
                    all_fit = False
                    logging.warn("TextFrame {} spill out still".format(name))
        if someframe_changed:
            return all_fit
        return None
