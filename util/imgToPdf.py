import libjpeg
from PIL import Image
import sys

#img = Image.fromarray(jpeg.decode(content))

with open(sys.argv[1], "rb") as f:
    img = Image.open(f)
    img.load()
with open(sys.argv[2], "wb") as f:
    img.save(f, 'PDF')
