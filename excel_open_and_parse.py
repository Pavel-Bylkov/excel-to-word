"""Работаем с пандой для получения словаря значений из иксель таблицы."""
from pandas import *
from docxtpl import DocxTemplate

xls = ExcelFile('AOOK2.xlsx')
data = xls.parse('static')

context = {}
f = 0
for n in data['keys']:
    context[n] = data['static'][f]
    f += 1

data = xls.parse('dynamic')
f = 0
for i in range(len(data.columns)):
    f += 1
    m = 0
    for k in data['Наименование']:
        context[k] = data[data.columns[f]][m]
        m += 1
    template = DocxTemplate(f'AOSR_TEMP.docx')
    template.render(context)
    template.save(f'AOSR_1_{i}.docx')
