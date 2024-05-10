import openpyxl
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import re

# Загружаем файл с необходимым названием из текущей директории
Tk().withdraw()
filename = askopenfilename().split('/')[-1]  # получаем название из добавленного файла
file = openpyxl.load_workbook('./' + filename) # открываем файл

# Выделяем страницу с необходимым названием
sheet = file.active
sheet.active = 1
column_f = sheet['F']

# Создаем первый список, куда добавляем диапозон, полностью охватывающий непустые ячейки.
first_list = []
for cell in range(len(column_f)):
    if isinstance(column_f[cell].value, str) and column_f[cell].value.find('prompt') != -1:
        if column_f[cell].value in first_list:
            continue
        else:
            first_list.append(re.sub('[А-Яа-яA-Z\W]', ' ', column_f[cell].value))

second_list = list(filter(lambda x: x.find('prompt') != -1, filter(None, ' '.join(first_list).split(' '))))

print(second_list)

# Создаем файл, куда записываем необходимые строки и список
with open('./validator.txt', 'w', encoding='utf-8') as file:
    file.write(f'def validator():\n\treturn {second_list}')
