#Probando

import openpyxl

book = openpyxl.load_workbook(r'C:\Users\Felipe\Desktop\EM Algorithm.xlsx')
sheet = book['Hoja2']

sheet['A3'] = 'Hola mundo'

book.save(r'C:\Users\Felipe\Desktop\EM Algorithm.xlsx')