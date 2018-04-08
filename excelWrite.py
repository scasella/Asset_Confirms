import xlsxwriter

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
