# Asset Confirms
This program extracts floats from an Excel file and searches local PDFs for matching floats using OCR.
1. User provides a directory ('dir') to the console.
2. Search dir for an .xlsx file
3. Once found, search .xlsx file for floats (dollars)
4. Search all PDFs in dir for matching floats
   - If not text-readable, use OCR to convert to text
5. Write matching floats and their locations to new .xlsx file. Also report non-matches.

### Kickoff
Finds first .xlsx file in directory, extracts floats, then breaks loop.
[excelRead()](#import-floats-from-xlsx-file)
```python
main() #Receives path by user in console then calls doExtract()

do_extract(path) #Path passed from main()
...
  for f_now in os.listdir(path):
    ext = os.path.splitext(f_now)[1]
    if ext.lower() in ['.xlsx']:
      float_dict = extract_excel(os.path.join(path, f_now))
      break
...
```
### Loop through each PDF
In doExtract()
```python
for f_now in os.listdir(path):
  ext = os.path.splitext(f_now)[1]
  if ext.lower() in ['.pdf']:
    print(path, f_now)
    doc_results = {}
    doc_results = convert_doc_pdf(os.path.join(path, f_now), float_dict) #Returns float matches in a PDF
```
Now in covertDocPDF()
[makeCV2(), doJPGPNG()](#convert-pdf-to-png-to-cv2-usable-for-ocr)
```python
convert_doc_pdf(path, float_dict)
...
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
```
Back in doExtract(), only do something if matches were found. Then return matches and original float list.
```python
if doc_results != None: #If any float matches were
  for key, val in doc_results.items():
    if key in end_dict:
      end_dict[key].append(val)
    else:
      end_dict[key] = [val]

print(len(end_dict))
return end_dict, float_dict
```
Finish by passing dicts to write new .xlsx file
[writeExcel()](#writes-output-to-new-xlsx-file)
```python
main()
...
  write_excel(var, end_dict, not_found)
  return
```
## Supporting Functions
#### Import floats from XLSX file
```python
def extract_excel(file):
  xl = pd.ExcelFile(file)
  df = xl.parse(xl.sheet_names[0])
  n_list = df.values.tolist()

  floats = {}
  for element in n_list:
    for i in element:
      if isinstance(i, float):
        floats[str(i)] = element
        break

  final_dict = {}
  for key, val in floats.items():
    temp_list = []
    for x in val:
      temp_list.append(str(x))
    final_dict[str(key)] = temp_list

  print('Distribution list imported')
  return final_dict
```

#### Convert PDF to PNG to CV2-usable for OCR
```python
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
```

#### Text-match OCR'd image and floats
```python
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
```
#### Writes output to new XLSX file
```python
def write_excel(path, f_dict, not_found):
  workbook = xlsxwriter.Workbook(path+'/Output.xlsx')
  worksheet = workbook.add_worksheet()

  format_h = workbook.add_format()
  format_h.set_bold()
  format_h.set_bg_color('#d9d9d9')
  format_h.set_align('vcenter')
  format_h.set_align('center')
  format_h.set_border(1)
  format_h.set_border_color('#000000')

  format_l = workbook.add_format()
  format_l.set_bold()
  format_l.set_font_color('blue')
  format_l.set_align('vcenter')

  format_a = workbook.add_format()
  format_a.set_align('vcenter')
  format_a.set_align('center')

  cols = ['Amount', 'Description', 'File']
  for ind, val in enumerate(cols):
    worksheet.write(1, ind+1, val, format_h)

  worksheet.set_column('A:A', 1.0)
  worksheet.set_column('B:B', 10)
  worksheet.set_column('C:C', 51)
  worksheet.set_column('D:D', 20)
  worksheet.set_row(0, 10)

  r_count = 2
  for key, val in f_dict.items():
    for inner_val in val:
      worksheet.write(r_count, 1, ('$'+key), format_a)
      worksheet.write(r_count, 2, inner_val[0])
      worksheet.write_url(r_count, 3, 'files\{0}'.format(inner_val[1].split(' ')[0]), string=inner_val[1], cell_format=format_l)
      r_count += 1

  r_count += 1
  worksheet.write(r_count, 1, "Not Found", format_h)
  r_count += 1
  for val in not_found:
    worksheet.write(r_count, 1, ('$'+val), format_a)
    r_count += 1

  workbook.close()
  return
```
