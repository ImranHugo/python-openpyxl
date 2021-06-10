from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
import datetime

def get_filename() -> str:
    basename = "output"
    suffix = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
    filename = "_".join([basename, suffix]) # e.g. 'mylogfile_120508_171442' 

    return filename


# create a workbook and prep the sheet layout
wb = Workbook()
ws = wb.active

column_lens = [4, 22, 33, 33, 22, 4, 35, 3]
for ch, c in zip('ABCDEFGH', column_lens):
    ws.column_dimensions[ch].width = c

row_heights = [(1, 65), (2, 65), (3, 25), (4, 33), (5, 27), (6, 27), (7, 20), (8, 36), (9, 27), (10, 27),
                (11, 27), (12, 27), (13, 20), (14, 33), (15, 27), (16, 27), (17, 27), (18, 27), (19, 27), (20, 18),
                (21, 30), (22, 55), (23, 30), (24, 70), (26, 20), (27, 30)]
for r, ch in row_heights:
    ws.row_dimensions[r].height = ch
    
# load up an image from disk
img = Image('image1.png')

# scale the image to fit your requirements
MAXHEIGHT = 84
h, w = img.height, img.width  # original size

ratio = MAXHEIGHT/h
h, w = h*ratio, w*ratio # scaled down/up size based on the ratio

p2e = pixels_to_EMU

# create the size profile of the image (scaled version)
size = XDRPositiveSize2D(p2e(w), p2e(h))

# Set image position on the top right of the Worksheet:
column = 8 # last column index (!!!starts from 0)
row = 0  # first row index (!!!starts from 0)
coloffset = -p2e(w) # we align the image with the right limit of the worksheet
marker = AnchorMarker(col=column, colOff=coloffset, row=row)
img.anchor = OneCellAnchor(_from=marker, ext=size)
ws.add_image(img) 

wb.save('{}.xlsx'.format(get_filename()))