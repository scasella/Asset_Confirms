import pandas as pd

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
