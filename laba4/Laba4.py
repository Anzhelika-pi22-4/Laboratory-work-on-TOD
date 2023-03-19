import xlwings as xw
from xlwings.constants import AutoFillType
import numpy as np
import pandas as pd
'Задачи для совместного разбора'
'Задача 1'

fail = xw.Book('себестоимостьА_в1.xlsx')
s = fail.sheets['Рецептура']
consumption = s.range('G7:O10').options(np.array).value
unit_price = s.range('G14:O14').options(np.array).value
rez = np.nan_to_num(consumption * unit_price).sum(axis=1)

'Задача 2'
'''
s.range('T7:T10').options(transpose=True).value = rez
'''
'Задача 3'
''''
s.range('T6').value = 'Себестоимость'
s.range('T4:T6').api.Merge()
s.range('T4:T6').color = (255, 0, 255)
'''''
'Задача 4'
''''
s.range('V7').formula = '=SUMPRODUCT(G7:O7, $G$14:$O$14)'
s.range('V7').api.Autofill(s.range('V7:V10').api, AutoFillType.xlFillDefault)
'''

'Лабораторная работа 7.1'
'Задача 1'

reviews_sample = pd.read_csv('reviews_sample.csv', sep = ',')
reviews_sample = reviews_sample.set_index('Unnamed: 0')
recipes_sample = pd.read_csv('recipes_sample.csv', sep = ',')
recipes_sample = recipes_sample[['id', 'name', 'minutes', 'submitted', 'description', 'n_ingredients']]


'Задача 2'

w1 = xw.Book()
w1.save('recipes.xlsx')
w1.sheets.add('Отзывы')
w1.sheets.add('Рецепты')
d1 = reviews_sample.sample(round(len(reviews_sample) * 0.05))
d2 = recipes_sample.sample(round(len(recipes_sample) * 0.05))
s1 = w1.sheets['Отзывы']
s2 = w1.sheets['Рецепты']
s1.range('A1').value = d1
s2.range('A1').value = d2
s1.range('A:A').api.Delete()
s2.range('A:A').api.Delete()

'Задача 3'
second_assign = (d2["minutes"] * 60).to_numpy()
s2.range('G1').value = 'second_assign'
s2.range('G2').options(transpose = True).value = second_assign

'Задача 4'
s2.range('H1').value = 'second_formula'
s2.range('H2').formula = '=C2*60'
'''''
s2.range('H2').api.Autofill(s2.range('H2:H1501').api, AutoFillType.xlFillDefault)
'''''

'Задача 5'
s1.range('A1:E1').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
s1.range('A1:E1').api.Font.Bold = True
s2.range('A1:I1').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
s2.range('A1:I1').api.Font.Bold = True

'Задача 6'
for i in s2.range(f'C2:C{len(d2) + 1}'):
    if i.value < 5:
        i.color = (0, 255, 0)
    elif (i.value >= 5 and i.value <= 10):
        i.color = (250, 255, 0)
    else:
        i.color = (255, 0, 0)

'Задача 7'
s2.range('I1').value = 'n_reviews'
s2.range('I2').formula = '=COUNTIF(Отзывы!$B$2:Отзывы!$B$6336, "="&Рецепты!A2)'
'''''
s2.range('I2').api.Autofill(s2.range('I2:I1501').api, AutoFillType.xlFillDefault)
'''''

'Лабораторная работа 7.2'
'Задача 8'
'''''
def validate(sheet):
    sheet.range('G2').formula = '=COUNTIF(Рецепты!$A$2:Рецепты!$A$1501,"="&B2)>0'
    sheet.range('G2').api.Autofill(sheet.range(f'G2:G{len(d1) + 1}').api, AutoFillType.xlFillDefault)
    for i in range(2, len(d1) + 2):
        rating = sheet.range(f'D{i}').value
        if (rating < 0 or rating > 5) or sheet.range(f'G{i}').value == False:
            sheet.range(f'A{i}:E{i}').color = (255, 0, 0)
    sheet.range('G:G').api.Delete()
validate(s1)
'''''












