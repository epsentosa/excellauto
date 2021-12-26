import openpyxl as xl
from utils import *
import os
import subprocess

user044 = '/home/user044/Documents/Python/'
ekoputra = '/home/ekoputra/Documents/Python/'

os.chdir(user044)

source_file = 'Subcon Master Data.xlsx'

wb = xl.load_workbook(source_file)
sheet = wb['ooh']
dest_sheet = wb.active

data_input = input_data(sheet)

n = 0
if len(data_input) == 0:
    print('No data Stored')
else:
    for cell in dest_sheet.iter_rows(min_row=4, max_row=3000):
        if cell[3].value == None:
            cell[3].value = data_input[n].value
            n += 1
            if n == len(data_input):
                break

del_rows(dest_sheet)

newfilename = input('Enter new file name: ')
wb.save(f'{newfilename}.xlsx')

print(f'{newfilename}.xlsx saved.')

opener = 'libreoffice'
subprocess.call([opener,f'{newfilename}.xlsx'])