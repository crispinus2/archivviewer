# CategoryModel.py

import io, struct, os, sys, logging
from PyQt5.QtCore import QAbstractListModel, QMutex, Qt
from PyQt5.QtGui import QBrush, QColor
from contextlib import contextmanager
from collections import OrderedDict
from itertools import islice

LOGGER = logging.getLogger(__name__)

class MOMask:
    NOTUSECATEGORY = 0b00001000
    DAYPROTOCOL = 0b00000100
    AUTOSEND = 0b00000001
    EMERGENCYSEND = 0b00000010

class ExceptionBytesIO(io.BytesIO):
    def __init__(self, buffer = None):
        io.BytesIO.__init__(self, buffer)
        
    def read(self, size):
        res = io.BytesIO.read(self, size)
        if len(res) < size:
            raise EOFError('End of stream reached')
        
        return res

def parse_memo_blob(blob):
    categories = {}
    offset = 0
    totallen = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    b1len = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2 + b1len
    b2len = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2+ b2len
    b3len = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2+ b3len
    b4len = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2 + b4len
    b5len = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2 + b5len
    #b6len = int.from_bytes(blob[offset:offset+2], 'little')
    #offset += 2 + b6len
    categoriesLenRelative = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    categoriesLastOffset = categoriesLenRelative + offset
    catCountLen = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    catCount = int.from_bytes(blob[offset:offset+catCountLen], 'little')
    offset += catCountLen
    while offset < categoriesLastOffset:
        entryLen = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        category = parse_memo_blob_category(blob[offset:offset+entryLen])
        offset += entryLen
        #if category['catType'] == "archive":
        if category['useCategory']:
            categories[category["id"]] = category
        
    return categories

def parse_memo_blob_category(blob):
    offset = 0
    name = None
    keycode = None
    red = None
    blue = None
    green = None
    unknown2 = None
    unknown3 = None
    unknown4 = None
    auftrag = None
    unknown6 = None
    key2 = None
    useCategory = True
    dayprotocol =  True
    autosend = False
    emergencysend = False
    externalfile = False
    catType = None
    
    b1len = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    b1 = blob[offset:offset+b1len]
    offset += b1len
    catidlen = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    catid = int.from_bytes(blob[offset:offset+catidlen], 'little')
    offset += catidlen
    kbtlen = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    kbt = blob[offset:offset+kbtlen-1].decode('cp1252')
    offset += kbtlen
    unklen5 = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    if unklen5 < 20:
        if unklen5 > 0:
            auftrag = blob[offset:offset+unklen5-1].decode('cp1252')
            offset += unklen5
        unklen3 = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        offset += unklen3
        unklen4 = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        infobytes = blob[offset:offset+unklen4]
        configbyte = 0b00000000
        if len(infobytes) > 0:
            configbyte = infobytes[0]
        if configbyte & MOMask.NOTUSECATEGORY:
            useCategory = False
        if configbyte & MOMask.DAYPROTOCOL:
            dayprotocol = False
        if configbyte & MOMask.AUTOSEND:
            autosend = True
        if configbyte & MOMask.EMERGENCYSEND:
            emergencysend = True
        offset += unklen4
        unklen6 = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        if unklen6 > 0:
            unknown6 = int.from_bytes(blob[offset:offset+unklen6], 'little')
            offset += unklen6
        namelen = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        name = blob[offset:offset+namelen-1].decode('cp1252')
        offset += namelen
        key2len = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        if key2len > 0:
            key2 = blob[offset:offset+key2len-1]
            offset += key2len
        else:
            key2 = None
        keycodelen = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        keycode = blob[offset:offset+keycodelen-1].decode('cp1252')
        offset += keycodelen
        if blob[offset:offset+2] == bytes.fromhex('0000'):
            offset += 2
            colorlen = int.from_bytes(blob[offset:offset+2], 'little')
            offset += 2
            if colorlen == 4:
                red = blob[offset]
                offset += 1
                green = blob[offset]
                offset += 1
                blue = blob[offset]
                offset += 2
            else:
                offset += colorlen
        else:
            red = None
            green = None
            blue = None       
    
    return { 'id': catid, 'krankenblatt': kbt,  'auftrag': auftrag,  'name': name, 'keycode': keycode, 'key2': key2, 'red': red, 'green': green, 
                'blue': blue, 'useCategory': useCategory, 'dayProtocol': dayprotocol, 'autosend': autosend, 'emergencysend': emergencysend}

class CategoryModel(QAbstractListModel):
    def __init__(self, con):
        super(CategoryModel, self).__init__()
        self._con = con
        self._mutex = QMutex()
        self._fullcategories = {}
        self._archivecategories = {}
        self.reloadCategories()

    @contextmanager
    def lock(self, msg = None):
        if msg is not None:
            #print("Lock request: {}".format(msg))
            pass
        self._mutex.lock()
        if msg is not None:
            #print("Lock acquired: {}".format(msg))
            pass
        try:
            yield
        except:
            raise
        finally:
            self._mutex.unlock()
            if msg is not None:
                #print("Lock released: {}".format(msg))
                pass
    
    def allCategories(self):
        with self.lock('allCategories'):
            return self._archivecategories
    
    def reloadCategories(self):
        cur = self._con.cursor()
        cur.execute("SELECT s.FMEMO, s.FBRIEFKATEGORIELISTE, s.FABLAGELISTE, s.FKATEGORIELISTE FROM MOSYSTEM s")
        for blobs in cur:
            self.beginResetModel()
            with self.lock("reloadCategories"):
                try:
                    self._fullcategories = parse_memo_blob(blobs[0])
                except:
                    dumpfile = os.sep.join([os.path.dirname(os.path.abspath(sys.argv[0])), "MO-Memo-Dump.hex"])
                    try:
                        with open(dumpfile, "wb") as f:
                            f.write(blobs[0])
                    except:
                        pass
                    raise
                archivecategories = {}
                filterprefixes = [ 'Bildarchiv', 'Externe Datei', 'Brief' ]
                for (catid, cat) in self._fullcategories.items():
                    LOGGER.debug("Splitting {}: {}".format(catid, cat["name"]))
                    try:
                        prefix, shortname = cat["name"].split(" - ", 1)
                        if prefix in filterprefixes:
                            archivecategories[catid] = { 'name': shortname, 'krankenblatt': cat['krankenblatt'] }
                    except ValueError:
                        pass
                    
                self._archivecategories = OrderedDict(sorted(archivecategories.items(), key=lambda item: item[1]['name']))
                
            self.endResetModel()
            break
        del cur
        
    def rowCount(self, _):
        with self.lock("rowCount"):
            return len(self._archivecategories)
    
    def categoryById(self, id):
        with self.lock("nameById"):
            return self._archivecategories[id]
    
    def idAtRow(self, row):
        with self.lock("idAtRow"):
            return list(self._archivecategories.keys())[row]
    
    def colorById(self, id):
        with self.lock("colorById"):
            fullcat = self._fullcategories[id]
            
        return { 'red': fullcat['red'], 'green': fullcat['green'], 'blue': fullcat['blue'] }
    
    def data(self, index, role):
        if role == Qt.DisplayRole:
            with self.lock("data / DisplayRole"):
                return '{name} ({krankenblatt})'.format(**list(self._archivecategories.values())[index.row()])
        elif role == Qt.DecorationRole:
            with self.lock("data / BackgroundRole"):
                try:
                    id = list(self._archivecategories.keys())[index.row()]
                    fullcat = self._fullcategories[id]
                    red = fullcat['red']
                    if red is not None:
                        green = fullcat['green']
                        blue = fullcat['blue']
                        return QColor.fromRgb(red, green, blue)
                    else:
                        return QColor.fromRgb(255, 255, 255, 0)
                except KeyError:
                    return QColor.fromRgb(255, 255, 255, 0)
