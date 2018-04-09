import xlsxwriter

def write_excel(path, f_dict, not_found, excel_file_path):
    workbook = xlsxwriter.Workbook(path+'/Output.xlsx')
    worksheet = workbook.add_worksheet()

    format_t = workbook.add_format()
    format_t.set_bold()
    format_t.set_align('vcenter')
    format_t.set_align('left')
    format_t.set_border(1)
    format_t.set_border_color('#000000')
    format_t.set_font_size(16)

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

    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 51)
    worksheet.set_column('C:C', 20)

    worksheet.merge_range('A1:C1','')
    worksheet.write_url('A1:C1', excel_file_path, format_t, string=('Distribution Confirmation: {0}'.format(excel_file_path.split('\\')[-1])))
    worksheet.set_row(0,28)

    cols = ['Amount', 'Description', 'File']
    for ind, val in enumerate(cols):
        worksheet.write(1, ind, val, format_h)

    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 51)
    worksheet.set_column('C:C', 20)

    r_count = 2
    for key, val in f_dict.items():
        for inner_val in val:
            worksheet.write(r_count, 0, ('$'+key), format_a)
            worksheet.write(r_count, 1, inner_val[0])
            worksheet.write_url(r_count, 2, 'files\{0}'.format(inner_val[1].split(' ')[0]), string=inner_val[1], cell_format=format_l)
            r_count += 1

    r_count += 1
    worksheet.write(r_count, 0, "Not Found", format_h)
    r_count += 1
    for val in not_found:
        worksheet.write(r_count, 0, ('$'+val), format_a)
        r_count += 1

    workbook.close()
    return
