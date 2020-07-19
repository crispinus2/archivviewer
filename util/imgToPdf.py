import libjpeg
from PIL import Image
import sys

with open(sys.argv[1], "rb") as f:
    img = Image.fromarray(libjpeg.decode(f.read(), 1))
with open(sys.argv[2], "wb") as f:
    img.save(f, 'PDF')
