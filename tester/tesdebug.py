import easyocr
from pathlib import Path
from pdf2image import convert_from_path
import numpy
ocrreader = easyocr.Reader(['en'])
images = convert_from_path(Path('pdftmp.pdf'))
imgcrop = images[0].crop(box = (180,750,750,900))
imgcrop.save(Path("pdftmp.png"))
# res = ocrreader.readtext(numpy.array(imgcrop)  , detail = 0)
res = ocrreader.readtext("pdftmp.png"  , detail = 0)


