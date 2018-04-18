import cv2
import pytesseract
import numpy as np
from wand.image import Image as Image
from wand.color import Color as Color
import PyPDF2
import re

class PDFObj:
    def __init__(self, path, float_dict):
        self.path = path
        self.float_dict = float_dict
        self.doc_results = {}

    def searchDists(self):
        self.doc_results = self.__handle_pdf(self.path, self.float_dict, self.doc_results)
        return self.doc_results

    def __handle_pdf(self, path, float_dict, doc_results):
        pdf_file_obj = open(path, 'rb')
        pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
        txt_check = pdf_reader.getPage(0).extractText()
        if txt_check != '':
            for i in range(pdf_reader.numPages):
                temp_text = pdf_reader.getPage(i).extractText().replace(',','').replace('\n','').replace('\r','')
                results = self.__find_text(temp_text, float_dict, [path, i+1, pdf_reader.numPages])
                doc_results = self.__diff_doc_results(results, doc_results)
                print('Scan {0} page {1}/{2} completed'.format(str(path).split('\\')[-1], i+1, pdf_reader.numPages))
        else:
            doc_results = self.__convert_doc_pdf(path, float_dict)
        return doc_results

    def __diff_doc_results(self, results, doc_results):
        for key, val in results.items():
            if key in doc_results:
                doc_results[key].append(val)
            else:
                doc_results[key] = val
        return doc_results

    def __convert_doc_pdf(self, path, float_dict):
        try:
            img = Image(filename=path, resolution=350)
            img_seq = img.sequence
        except:
            print('Corrupt File - {0}'.format(path))
            return

        doc_results = {}
        for i in range(len(img_seq)):
            cv_ready = self.__make_cv2(Image(image=img_seq[i]))
            text = self.__do_jpg_png(cv_ready)
            results = self.__find_text(text, float_dict, [path, i+1, len(img_seq)]) #Find matches between float_dict and given text
            doc_results = self.__diff_doc_results(results, doc_results)
            print('Scan {0} page {1}/{2} completed'.format(str(path).split('\\')[-1], i+1, len(img_seq)))
        return doc_results

    def __make_cv2(self, img):
        img.convert("png")
        img.background_color = Color('white')
        img.format = 'tif'
        img.alpha_channel = 'remove'
        #Convert PNG to cv2-usable PNG
        img_buffer = np.asarray(bytearray(img.make_blob()), dtype=np.uint8)
        retval = cv2.imdecode(img_buffer, 0)
        return retval

    def __do_jpg_png(self, cvF, ssn_re=''):
        pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract'
        gray = cv2.bitwise_not(cvF)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
        #thresh = cv2.adaptiveThreshold(cvF, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 3, 1)
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

    def __find_text(self, text, float_dict, docLoc):
        regx_term = re.compile('[A-Z].*\s[A-Z].*')
        found_dict = {}
        for key, val in float_dict.items():
            if key in text:
                name = 'No name'
                for i in val:
                    if re.search(regx_term, i) != None:
                        name = i
                        break
                doc_str = str(docLoc[0]).split('\\')[-1]
                page_str = 'page {0}/{1}'.format(docLoc[1], docLoc[2])
                final_loc = doc_str+' - '+page_str
                found_dict[key] = [name, final_loc]
        return found_dict
