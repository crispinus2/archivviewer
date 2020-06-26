#!/usr/bin/python3

import sys

class MOMask:
    NOTUSECATEGORY = 0b00001000
    DAYPROTOCOL = 0b00000100
    AUTOSEND = 0b00000001
    EMERGENCYSEND = 0b00000010
    EXTERNALFILE = 0b00000100
    ARCHIVE = 0b00001010

def parse_blob(blob):
    categories = []
    offset = 0
    totallen = int.from_bytes(blob[offset:offset+2], 'little')
    print("Total length: {}".format(totallen))
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
    print("Categories Last Offset: {}/{}".format(categoriesLenRelative, categoriesLastOffset))
    catCountLen = int.from_bytes(blob[offset:offset+2], 'little')
    print ("Length of Cat Count: {}".format(catCountLen))
    offset += 2
    catCount = int.from_bytes(blob[offset:offset+catCountLen], 'little')
    print("Number of Categories: {}".format(catCount))
    offset += catCountLen
    while offset < categoriesLastOffset:
        entryLen = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        _, category = parse_blob_category(blob[offset:offset+entryLen], 0)
        offset += entryLen
        #if category['catType'] == "archive":
        categories.append(category)
        print("Read category: {}".format(category))
        
    return categories

def parse_blob_category(blob, offset):
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
    
    return (offset,
             { 'catid': catid, 'krankenblatt': kbt,  'auftrag': auftrag,  'name': name, 'keycode': keycode, 'key2': key2, 'red': red, 'green': green, 
                'blue': blue, 'useCategory': useCategory, 'dayProtocol': dayprotocol, 'autosend': autosend, 'emergencysend': emergencysend})

def parse_blob_briefe(blob):
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
        result = parseBlobEntry(blob[offset:offset+entryLength])
        entries[result['categoryId']] = result['name']
        offset += entryLength

if __name__ == "__main__":
    with open(sys.argv[1], "rb") as f:
        blob = f.read()
        parse_blob(blob)
