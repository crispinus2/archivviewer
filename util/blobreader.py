#!/usr/bin/python3

import sys

class MOMask:
    NOTUSECATEGORY = 0b00001000
    DAYPROTOCOL = 0b00000100

def parse_blob(blob):
    categories = []
    offset = 48
    while offset < len(blob):
        offset, category = parse_blob_category(blob, offset)
        if category['catid'] > 0:
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
    catid = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
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
                unknown3 = int.from_bytes(blob[offset:offset+2], 'little')
                offset += 2
            else:
                offset += colorlen
        else:
            unknown3 = int.from_bytes(blob[offset:offset+2], 'little')
            offset += 2
            red = None
            green = None
            blue = None       
    unklen2 = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    unknown2 = blob[offset:offset+unklen2]
    offset += unklen2
    offset += 2
    return (offset,
             { 'catid': catid, 'krankenblatt': kbt,  'auftrag': auftrag,  'name': name, 'keycode': keycode, 'key2': key2, 'red': red, 'green': green, 
                'blue': blue, 'unknown2': unknown2, 'unknown3': unknown3, 'unknown4': unknown4, 'unknown6': unknown6 })

if __name__ == "__main__":
    with open(sys.argv[1], "rb") as f:
        blob = f.read()
        parse_blob(blob)
