import PyPDF2
import pytesseract
import os,cv2
import numpy as np
from wand.image import Image as Image
from wand.color import Color as Color

from excelRead import *
from excelWrite import *

def main(var):
    endDict,floatDict = doExtract(var)
    if len(floatDict) == 0: return

    notFound = []
    for key in floatDict.keys():
        if key not in endDict:
            notFound.append(key)

    writeExcel(var,endDict,notFound)

def doExtract(path):
    endDict = {}
    floatDict = {}

    for f in os.listdir(path):
        ext = os.path.splitext(f)[1]
        if ext.lower() in ['.xlsx']:
            floatDict = extractExcel(os.path.join(path,f))
            break
    if len(floatDict) == 0: return endDict,floatDict
    #pdfFileObj = open(file, 'rb')
    #pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    #allTxt = ""

    #for i in range(pdfReader.numPages):
    #    allTxt+=pdfReader.getPage(i).extractText()

    #if repr(allTxt)=='':
    #    text = convertPDF(file)
    for f in os.listdir(path)[:2]:
        ext = os.path.splitext(f)[1]
        if ext.lower() in ['.pdf']:
            print(path,f)
            docResults = {}
            docResults = convertDocPDF(os.path.join(path,f),floatDict)
            if docResults != None:
                for key,val in docResults.items():
                    if key in endDict:
                        endDict[key].append(val)
                    else:
                        endDict[key] = [val]
    print(len(endDict))
    return endDict,floatDict

def convertDocPDF(path,floatDict):
    img_buffer = None
    docResults = {}
    try:
        img = Image(filename=path, resolution=350)
    except:
        print('Corrupt File - {0}'.format(path))
        return

    imgSeq = img.sequence

    for i in range(len(imgSeq)):
        cvReady = makeCV2(Image(image=imgSeq[i]))
        text = doJPGPNG(cvReady)
        results = findText(text,floatDict,[path,i+1,len(imgSeq)])
        for key,val in results.items():
            if key in docResults:
                docResults[key].append(val)
            else:
                docResults[key] = val
        print('Scan {0} page {1}/{2} completed'.format(str(path).split('\\')[-1],i+1,len(imgSeq)))
    return docResults


def findText(text,floatDict,docLoc):
    foundDict = {}

    for key in floatDict.keys():
        if key in text:
            ind = text.index(key)
            bText = text[:ind].split('\n')[-1]
            aText = text[ind:].split('\n')[0]
            finalT = [bText+aText]
            docStr = str(docLoc[0]).split('\\')[-1]
            pageStr = 'page {0}/{1}'.format(docLoc[1],docLoc[2])
            finalT.append(docStr+' - '+pageStr)
            foundDict[key] = finalT
    return foundDict


def makeCV2(img):
    img.convert("png")
    img.background_color = Color('white')
    img.format = 'tif'
    img.alpha_channel = 'remove'

    #Convert PNG to cv2-usable PNG
    img_buffer=np.asarray(bytearray(img.make_blob()), dtype=np.uint8)
    retval = cv2.imdecode(img_buffer,0)
    return retval


def doJPGPNG(cvF,ssnRE=''):
    pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract'

    gray = cv2.bitwise_not(cvF)
    #thresh = cv2.adaptiveThreshold(cvF, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 3, 1)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    #kernel = np.ones((1,1),np.uint8)
    #thresh = cv2.erode(thresh,kernel,iterations = 1)
    #blur = cv2.GaussianBlur(thresh,(1,1),0)
    #thresh = cv2.addWeighted(blur,1.5,thresh,-0.5,0)
    #im2 = cv2.resize(thresh, (int(thresh.shape[0]*0.25), int(thresh.shape[1]*0.25)))
    #cv2.imshow('img',im2)
    #cv2.waitKey(0)
    #cv2.destroyAllWindows()

    #Tesseract find text from img
    #final = ssnRE.findall(pytesseract.image_to_string(thresh).replace(' ', ''))
    #final = formatSSN(final)
    final = pytesseract.image_to_string(thresh)
    #print(final)
    #print(final)
    return final

var = input("Please enter a folder path: ")
print("Scanning...")
main(var)
