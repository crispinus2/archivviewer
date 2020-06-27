# CategoryModel.py

import io
from PyQt5.QtCore import QAbstractListModel, QMutex, Qt
from PyQt5.QtGui import QBrush, QColor
from contextlib import contextmanager
from collections import OrderedDict
from itertools import islice

class MOMask:
    NOTUSECATEGORY = 0b00001000
    DAYPROTOCOL = 0b00000100
    AUTOSEND = 0b00000001
    EMERGENCYSEND = 0b00000010

def parseBlobs(blobs):
    entries = {}
    for blob in blobs:
        offset = 0
        totalLength =  int.from_bytes(blob[offset:offset+2], 'little')
        offset += 4
        entryCount = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        
        while offset < len(blob):
            entryLength = int.from_bytes(blob[offset:offset+2], 'little')
            offset += 2
            result = parseBlobEntry(blob[offset:offset+entryLength])
            entries[result['categoryId']] = result['name']
            offset += entryLength
    
    return entries

def parseBlobEntry(blob):
    offset = 6
    name_length = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    name = blob[offset:offset+name_length-1].decode('cp1252')
    offset += name_length
    catid = int.from_bytes(blob[-2:],  'little')
    
    return { 'name': name, 'categoryId': catid }

def parse_briefe_blob(blob):
    entries = []
    offset = 0
    totalLength =  int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    entryCountLength = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    entryCount = int.from_bytes(blob[offset:offset+entryCountLength], 'little')
    offset += entryCountLength
        
    while offset < len(blob):
        entryLength = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        result = parse_briefe_entry(blob[offset:offset+entryLength])
        entries.append(result)
        offset += entryLength
    
    return entries

def parse_briefe_entry(blob):
    stream = io.BytesIO(blob)
    b1len = int.from_bytes(stream.read(2), 'little')
    stream.read(b1len)
    catid = None
    namelen = int.from_bytes(stream.read(2), 'little')
    name = stream.read(namelen)[:-1].decode('cp1252')
    keycodelen = int.from_bytes(stream.read(2), 'little')
    if keycodelen == 0:
        stream.read(4)
        keycodelen = int.from_bytes(stream.read(2), 'little')
        keycode = stream.read(keycodelen)[:-1].decode('cp1252')
        catidlen = int.from_bytes(stream.read(2), 'little')
        catid = int.from_bytes(stream.read(catidlen), 'little')
    else:
        keycode = stream.read(keycodelen)[:-1].decode('cp1252')
    if len(keycode) < 2:
        keycode = ''.join(['q', keycode])
        
    return { 'categoryId': catid, 'name': name, 'keycode': keycode }
    
        

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
        
    def reloadCategories(self):
        cur = self._con.cursor()
        cur.execute("SELECT s.FMEMO, s.FBRIEFKATEGORIELISTE, s.FKATEGORIELISTE, s.FABLAGELISTE FROM MOSYSTEM s")
        for blobs in cur:
            self.beginResetModel()
            with self.lock("reloadCategories"):
                self._fullcategories = parse_memo_blob(blobs[0])
                briefcategories = parse_briefe_blob(blobs[1])
                archivecategories = parseBlobs(blobs[2:])
                
                for bc in briefcategories:
                    if bc['categoryId'] is not None:
                        archivecategories[bc['categoryId']] = bc['name']
                    else:
                        for fc in self._fullcategories.values():
                            if fc['krankenblatt'].lower() == bc['keycode'].lower():
                                archivecategories[fc['id']] = bc['name']
                                break
                archivecategories = { k: { 'name': v, 'krankenblatt': self._fullcategories[k]['krankenblatt'] } if k in self._fullcategories else { 'name': v, 'krankenblatt': v } 
                                     for (k, v) in archivecategories.items() }
                
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
        elif role == Qt.BackgroundRole:
            with self.lock("data / BackgroundRole"):
                try:
                    id = list(self._archivecategories.keys())[index.row()]
                    fullcat = self._fullcategories[id]
                    red = fullcat['red']
                    if red is not None:
                        green = fullcat['green']
                        blue = fullcat['blue']
                        return QBrush(QColor.fromRgb(red, green, blue))
                except KeyError:
                    pass
