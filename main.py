import os

from pdfClass import PDFObj
from excelRead import extract_excel
from excelWrite import write_excel
import re

import warnings
warnings.simplefilter("ignore") #To surpress PDF-read 'superfluous whitespace' warning

def main(var):
    end_dict, float_dict, excel_file_path = do_extract(var)
    if any(float_dict) == 0:
        print('No values found in Excel file')
        return

    not_found = []
    for key, val in float_dict.items():
        if key not in end_dict:
            item = key
            name = 'No name'
            regx_term = re.compile('[A-Z].*\s[A-Z].*')
            for i in val:
                if re.search(regx_term, i) != None:
                    name = i
                    break
            not_found.append([name, item])

    write_excel(var, end_dict, not_found, excel_file_path)
    return

def do_extract(path):
    end_dict = {}
    float_dict = {}

    #Get floats from first Excel file then break
    for f_now in os.listdir(path):
        ext = os.path.splitext(f_now)[1]
        if ext.lower() in ['.xlsx']:
            float_dict = extract_excel(os.path.join(path, f_now))
            excel_file_path = str(os.path.join(path, f_now))
            break

    #Return if empty float dict (no Excel floats)
    if any(float_dict) == 0:
        return end_dict, float_dict, excel_file_path

    pdf_obj_dict = {}
    for f_now in os.listdir(path):
        ext = os.path.splitext(f_now)[1]
        if ext.lower() in ['.pdf']:
            print(path, f_now)
            #Adds to pdf_obj_dict[path]: PDFObj
            pdf_obj_dict[str(os.path.join(path, f_now)).split('\\')[-1].split('.')[0]] = PDFObj(os.path.join(path, f_now), float_dict)

    for key, pdf_class in pdf_obj_dict.items():
        doc_results = pdf_class.searchDists()
        if doc_results != None: #If any float matches were found
            for key, val in doc_results.items():
                if key in end_dict:
                    #If amount already in end_dict
                    end_dict[key].append(val)
                else:
                    #New amount match found, val is location
                    end_dict[key] = [val]

    print(len(end_dict))
    return end_dict, float_dict, excel_file_path

while True:
    var = input("Please enter a folder path or help: ")
    if var.lower().replace(' ','') == 'help':
        print('Place the distribution list Excel file in a folder with all asset statement PDFs you want to scan. Then paste the folder path below and press Enter.')
    else:
        break
print("Scanning...")
main(var)
print('Scanning completed.')