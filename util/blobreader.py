#!/usr/bin/python3

import sys

def parse_blob(blob):
    categories = []
    offset = 48
    while offset < len(blob):
        offset, category = parse_blob_category(blob, offset)
        categories.append(category)
        print("Read category: {}".format(category))
        
    return categories

def parse_blob_category(blob, offset):
    
    catid = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    kbtlen = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    kbt = blob[offset:offset+kbtlen-1].decode('cp1252')
    offset += kbtlen
    offset += 2
    unklen3 = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    offset += unklen3
    unklen4 = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    offset += unklen4
    offset += 2
    namelen = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    name = blob[offset:offset+namelen-1].decode('cp1252')
    offset += namelen
    offset += 3
    keycodelen = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    keycode = blob[offset:offset+keycodelen-1].decode('cp1252')
    offset += keycodelen
    offset += 2
    colorlen = int.from_bytes(blob[offset:offset+2], 'little')
    print("Colorlen is {}".format(colorlen))
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
        red = None
        green = None
        blue = None
    unknown1 = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    unklen2 = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    unknown2 = blob[offset:offset+unklen2]
    offset += unklen2
    offset += 2
    return (offset,
             { 'catid': catid, 'krankenblatt': kbt,  'name': name, 'keycode': keycode, 'red': red, 'green': green, 'blue': blue, 'unknown1': unknown1, 'unknown2': unknown2 })

if __name__ == "__main__":
    with open(sys.argv[1], "rb") as f:
        blob = f.read()
        parse_blob(blob)
