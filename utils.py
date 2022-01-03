from openpyxl.utils import get_column_letter
from openpyxl.styles.borders import Border, Side
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from typing import Any

#contoh hasil konvert (21/E/SB/C/GIOD/00767 - B to GIOD/767/B)
def convert_cmt(cell_data, number='0'):
    index_second_char = cell_data.value.find(number)
    last_char = cell_data.value[-1::1]
    if last_char.isdigit():
        last_char = 'A'

    first_data_value = cell_data.value[10:15]
    second_data_value = int(cell_data.value[index_second_char:index_second_char+5])
    return first_data_value+str(second_data_value)+"/"+last_char

#Menghapus kolom yang ditentukan
#list tuple sesuai preferensi sendiri !!!
def del_columns(worksheet):
    tuple_for_cols = ([3,2],[5,1],[6,2],[9,3],[10,4],[12,7],[13,1],[14,1],[15,2],[16,6],[17,13])
    
    for col in tuple_for_cols:
        worksheet.delete_cols(col[0],col[1])

#Unhide semua kolom
def unhide_col(worksheet):
    len_col = worksheet.max_column
    for cols in range(1,len_col+1):
        col = get_column_letter(cols)
        worksheet.column_dimensions[col].hidden = False

#memberi nomor sampai max value colomn
def add_number(worksheet):
    len_col = worksheet.max_column
    for rows in range(1):
        for cols in range(1,len_col+1):
            worksheet.cell(2,cols).value = cols

#Membuat auto width kolom
def auto_fit(worksheet):
    len_col = worksheet.max_column
    len_row = worksheet.max_row
    for data_cols in range(1,len_col+1):
        column_width = 0
        letter = get_column_letter(data_cols)
        for data in range(1,len_row+1):
            if worksheet[letter+str(data)].value == None:
                continue
            elif column_width < len(str(worksheet[letter+str(data)].value)):
                column_width = len(str(worksheet[letter+str(data)].value))
            if column_width > 25:
                column_width = 25
        worksheet.column_dimensions[letter].width = column_width+2

#Mengapus baris yang kosong
def del_rows(sheet):
    len_row = sheet.max_row
    index_row = []
    # loop each row in column
    for index in range(6, len_row+1):
        # define emptiness of cell
        val = sheet.cell(index,1)
        if val.value is None or val.value[:1].isdigit() == False:
            # collect indexes of rows
            index_row.append(index)

    # loop each index value
    for row_del in range(len(index_row)):
        sheet.delete_rows(idx=index_row[row_del], amount=1)
        # exclude offset of rows through each iteration
        index_row = list(map(lambda k: k - 1, index_row))


#Set border untuk Cell
def set_border(style_desc):
    return Border(left=Side(style=style_desc), 
                        right=Side(style=style_desc), 
                        top=Side(style=style_desc), 
                        bottom=Side(style=style_desc))

#Auto border range yang ditentukan    
def auto_border(sheet,end_row,end_col,style='thin',start_row=1,start_col=1):
    for row in range(start_row,end_row+1):
        for col in range(start_col,end_col+1):
            sheet.cell(row,col).border = set_border(style)

#Ambil copy data row/column untuk disimpan di list database
def take_data(sheet,range_copy: Any='Enter Row Number or Collumn Letter'):
    data_copy = []
    if isinstance(range_copy,int):
        source_row = sheet.iter_rows(min_row=range_copy, max_row=range_copy)
        for row in source_row:
            for col in row:
                data_copy.append(col.value)
    else:
        range_copy = column_index_from_string(range_copy)
        source_col = sheet.iter_cols(min_col=range_copy, max_col=range_copy)
        for col in source_col:
            for row in col:
                data_copy.append(row.value)
    return data_copy

#Paste data dari database yang sudah dicopy
def put_data(sheet,list_data,paste_as: Any='Type row or col',cell: Any='Cell Destination'):
    cell = coordinate_from_string(cell)
    row_number = cell[1]
    column_letter = column_index_from_string(cell[0])
    if paste_as.lower() == 'row':
        dest_row = sheet.iter_rows(min_row=row_number, max_row=row_number,min_col=column_letter, max_col=len(list_data)+column_letter-1)
        for row in dest_row:
            i = 0
            for cell in row:
                cell.value = list_data[i]
                i += 1
    elif paste_as.lower() == 'col':
        dest_col = sheet.iter_cols(min_col=column_letter, max_col=column_letter,min_row= row_number, max_row=len(list_data)+row_number-1)
        for col in dest_col:
            i = 0
            for cell in col:
                cell.value = list_data[i]
                i += 1

#Fungsi menerima input data dari User -> cek database -> menyimpan di list
def input_data(sheet):
    data_input = []
    while True:
        input_value = input('Enter CMT No : ')
        if input_value.lower() == 'finish':
            print('Finish input')
            break
        #n = []
        for cell in sheet.iter_rows(min_row=6,max_col=2):
            cell_value = cell[1].value
            if input_value.upper() in cell_value:
                data_input.append(cell[1])
                print(cell_value,' Added')
            
            #Fungi dibawah menambahkan pesan apabila tidak ada dalam 1x loop, tetapi dengan  efek samping lama
            '''
            else:
                n += 1
                if n == len(list(sheet.iter_rows(min_row=6,max_col=1))):
                    print('No Data, pls enter correct CMT') '''
    return data_input