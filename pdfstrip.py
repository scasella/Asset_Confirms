import PyPDF2

from wand.image import Image as Image2
from wand.color import Color as Color2

pdf1File = open('stmt2b.pdf', 'rb')
pdf1Reader = PyPDF2.PdfFileReader(pdf1File)

pdfWriter = PyPDF2.PdfFileWriter()

for pageNum in range(pdf1Reader.numPages):
    pageObj = pdf1Reader.getPage(pageNum)
    pdfWriter.addPage(pageObj)

pdfOutputFile = open('better.pdf', 'wb')
pdfWriter.write(pdfOutputFile)
pdfOutputFile.close()
pdf1File.close()

with Image2(filename='better.pdf', resolution=300) as img:
    print('done')
    img.convert("png")
    img.background_color = Color2('white')
    img.format = 'tif'
    img.alpha_channel = False
