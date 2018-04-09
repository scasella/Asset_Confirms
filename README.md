# Distribution Confirmation
User pastes a directory ('dir') in console.
1. Search dir for an .xlsx file
2. Once found, search .xlsx file for floats (dollars)
3. Search all PDFs in dir for matching floats
   - If not text-readable, use OCR to convert to text
4. Write matching floats and their locations to new .xlsx file. Also report non-matches.

### Kickoff
Finds first .xlsx file in directory, extracts floats, the breaks loop.
[excelRead()](#imports-floats-from-xlsx-file)
```python
main() #Receives path by user in console then calls doExtract()

doExtract(path) #Path passed from main()
...
  for f in os.listdir(path):
    ext = os.path.splitext(f)[1]
    if ext.lower() in ['.xlsx']:
      floatDict = extractExcel(os.path.join(path,f))
      break
...
```
### Loop through each PDF
In doExtract()
```python
for f in os.listdir(path):
  ext = os.path.splitext(f)[1]
  if ext.lower() in ['.pdf']:
    print(path,f)
    docResults = {}
    docResults = convertDocPDF(os.path.join(path,f),floatDict) #Returns float matches in a PDF
```
Now in covertDocPDF()
[makeCV2(), doJPGPNG()](#converts-pdf-to-png-to-cv2-usable)
```python
convertDocPDF(path,floatDict)
...
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
    results = findText(text,floatDict,[path,i+1,len(imgSeq)]) #Find matches between floatDict and given text
    for key,val in results.items():
      if key in docResults:
        docResults[key].append(val)
      else:
        docResults[key] = val
    print('Scan {0} page {1}/{2} completed'.format(str(path).split('\\')[-1],i+1,len(imgSeq)))
  return docResults
```
Back in doExtract(), only do something if matches were found. Then return matches and original float list.
```python
if docResults != None: #If any float matches were found in PDFs
  for key,val in docResults.items():
    if key in endDict:
      endDict[key].append(val)
    else:
      endDict[key] = [val]

print(len(endDict))
return endDict,floatDict
```
Finish by passing dicts to write new .xlsx file
[writeExcel()](#writes-output-xlsx-file)
```python
main()
...
  writeExcel(var,endDict,notFound)
  return
```
## Supporting Functions
#### Imports floats from XLSX file
```python
def extractExcel(file):
  xl = pd.ExcelFile(file)
  df = xl.parse(xl.sheet_names[0])
  nList = df.values.tolist()

  floats = {}
  for element in nList:
    for i in element:
      if isinstance(i, float):
        floats[str(i)] = element
        break

  finalDict = {}
  for key,val in floats.items():
    tempList = []
    for x in val:
      tempList.append(str(x))
    finalDict[str(key)] = tempList

  print('Distribution list imported')
  return finalDict
```

#### Converts PDF to PNG to CV2-usable
```python
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
```

#### Text-matching
```python
def findText(text,floatDict,docLoc):
  foundDict = {}

  for key in floatDict.keys():
    if key in text:
      ind = text.index(key)
      bText = text[:ind].split('\n')[-1]
      bInd = 2
      while len(bText) < 12:
        bText = text[:ind].split('\n')[-bInd]+' '+bText
        bInd+=1

      aText = text[ind:].split('\n')[0]
      finalT = [bText+aText]
      docStr = str(docLoc[0]).split('\\')[-1]
      pageStr = 'page {0}/{1}'.format(docLoc[1],docLoc[2])
      finalT.append(docStr+' - '+pageStr)
      foundDict[key] = finalT
  return foundDict
```
#### Writes output XLSX file
```python
def writeExcel(path,fDict,notFound):
  workbook = xlsxwriter.Workbook(path+'/Output.xlsx')
  worksheet = workbook.add_worksheet()

  formatH = workbook.add_format()
  formatH.set_bold()
  formatH.set_bg_color('#d9d9d9')
  formatH.set_align('vcenter')
  formatH.set_align('center')
  formatH.set_border(1)
  formatH.set_border_color('#000000')

  formatL = workbook.add_format()
  formatL.set_bold()
  formatL.set_font_color('blue')
  formatL.set_align('vcenter')

  formatA = workbook.add_format()
  formatA.set_align('vcenter')
  formatA.set_align('center')

  cols = ['Amount','Description','File']
  for ind,val in enumerate(cols):
    worksheet.write(1,ind+1,val,formatH)

  worksheet.set_column('A:A',1.0)
  worksheet.set_column('B:B',10)
  worksheet.set_column('C:C',51)
  worksheet.set_column('D:D',20)
  worksheet.set_row(0,10)

  rCount = 2
  for key,val in fDict.items():
    for innerVal in val:
      worksheet.write(rCount,1,('$'+key),formatA)
      worksheet.write(rCount,2,innerVal[0])
      worksheet.write_url(rCount,3, 'files\{0}'.format(innerVal[1].split(' ')[0]), string=innerVal[1], cell_format=formatL)
      rCount+=1

  rCount+=1
  worksheet.write(rCount,1,"Not Found",formatH)
  rCount+=1
  for val in notFound:
    worksheet.write(rCount,1,('$'+val),formatA)
    rCount+=1

  workbook.close()
  return
```
