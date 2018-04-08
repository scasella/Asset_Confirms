import pandas as pd

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
