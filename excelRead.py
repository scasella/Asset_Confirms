import pandas as pd

def extract_excel(file):
    xl = pd.ExcelFile(file)
    df = xl.parse(xl.sheet_names[0])
    df.fillna('', inplace=True)
    n_list = df.values.tolist()

    floats = {}
    for row in n_list:
        for cell in row:
            try:
                float(str(cell).replace(',',''))
                print(cell)
                #{float:[all row items]}
                floats[str(cell)] = [str(x) for x in row]
                break
            except:
                pass
    #floats = {'amt': [data in row (name, amt)]}
    print('Distribution list imported')
    return floats
