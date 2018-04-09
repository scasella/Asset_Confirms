import os
import cv2
import pytesseract
import numpy as np
from wand.image import Image as Image
from wand.color import Color as Color

from excelRead import extract_excel
from excelWrite import write_excel

def main(var):
    end_dict, float_dict = do_extract(var)
    if any(float_dict) == 0:
        return

    not_found = []
    for key in float_dict.keys():
        if key not in end_dict:
            not_found.append(key)

    write_excel(var, end_dict, not_found)
    return

def do_extract(path):
    end_dict = {}
    float_dict = {}

    for f_now in os.listdir(path):
        ext = os.path.splitext(f_now)[1]
        if ext.lower() in ['.xlsx']:
            float_dict = extract_excel(os.path.join(path, f_now))
            break
    if any(float_dict) == 0:
        return end_dict, float_dict
    #pdfFileObj = open(file, 'rb')
    #pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    #allTxt = ""

    #for i in range(pdfReader.numPages):
    #    allTxt+=pdfReader.getPage(i).extractText()

    #if repr(allTxt)=='':
    #    text = convertPDF(file)
    for f_now in os.listdir(path):
        ext = os.path.splitext(f_now)[1]
        if ext.lower() in ['.pdf']:
            print(path, f_now)
            doc_results = {}
            doc_results = convert_doc_pdf(os.path.join(path, f_now), float_dict) #Returns float matches in a PDF
            if doc_results != None: #If any float matches were
                for key, val in doc_results.items():
                    if key in end_dict:
                        end_dict[key].append(val)
                    else:
                        end_dict[key] = [val]
    print(len(end_dict))
    return end_dict, float_dict

def convert_doc_pdf(path, float_dict):
    doc_results = {}
    try:
        img = Image(filename=path, resolution=350)
    except:
        print('Corrupt File - {0}'.format(path))
        return

    img_seq = img.sequence

    for i in range(len(img_seq)):
        cv_ready = make_cv2(Image(image=img_seq[i]))
        text = do_jpg_png(cv_ready)
        results = find_text(text, float_dict, [path, i+1, len(img_seq)]) #Find matches between float_dict and given text
        for key, val in results.items():
            if key in doc_results:
                doc_results[key].append(val)
            else:
                doc_results[key] = val
        print('Scan {0} page {1}/{2} completed'.format(str(path).split('\\')[-1], i+1, len(img_seq)))
    return doc_results


def find_text(text, float_dict, docLoc):
    found_dict = {}

    for key in float_dict.keys():
        if key in text:
            ind = text.index(key)
            b_text = text[:ind].split('\n')[-1]
            b_ind = 2
            while len(b_text) < 12:
                b_text = text[:ind].split('\n')[-b_ind]+' '+b_text
                b_ind += 1

            a_text = text[ind:].split('\n')[0]
            final_t = [b_text+a_text]
            doc_str = str(docLoc[0]).split('\\')[-1]
            page_str = 'page {0}/{1}'.format(docLoc[1], docLoc[2])
            final_t.append(doc_str+' - '+page_str)
            found_dict[key] = final_t
    return found_dict


def make_cv2(img):
    img.convert("png")
    img.background_color = Color('white')
    img.format = 'tif'
    img.alpha_channel = 'remove'

    #Convert PNG to cv2-usable PNG
    img_buffer = np.asarray(bytearray(img.make_blob()), dtype=np.uint8)
    retval = cv2.imdecode(img_buffer, 0)
    return retval


def do_jpg_png(cvF, ssn_re=''):
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
    #final = ssn_re.findall(pytesseract.image_to_string(thresh).replace(' ', ''))
    #final = formatSSN(final)
    final = pytesseract.image_to_string(thresh)
    return final

var = input("Please enter a folder path: ")
print("Scanning...")
main(var)
