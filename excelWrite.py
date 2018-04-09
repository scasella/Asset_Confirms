import xlsxwriter

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
