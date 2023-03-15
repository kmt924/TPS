import pandas as pd
import re

df0 = pd.read_excel('ЦА_ТО_Сведения_загрузки_данных.xlsx',
                    sheet_name='TDSheet')

searchfor = ["Номер приказа:", "дата приказа:", "Не найдено соответствие для реквизита:",
             "'Сотрудник' идентификатор:", "'Приказ о назначении сотрудника' идентификатор:"]
#searchfor = ['Не найдено соответствие для реквизита:', 'Приказ о назначении сотрудника',
#             'На дату кадрового приказа не найден действующий сотрудник']

#################         ПОИСК нескольких вхождений по "ИЛИ"

df1_ = df0[df0['Ошибка'].str.contains('|'.join(searchfor), na= False)]
#df1_ = df0[df0['Ошибка'].str.contains('|'.join(map(re.escape, searchfor)), na= False)]

#################         ПОИСК нескольких вхождений по "И"
df1 =df0[(df0['Ошибка'].str.contains("Номер приказа:")) &
         (df0['Ошибка'].str.contains("дата приказа:")) &
         (df0['Ошибка'].str.contains("Не найдено соответствие для реквизита:")) &
         (df0['Ошибка'].str.contains("'Сотрудник' идентификатор:")) &
         (df0['Ошибка'].str.contains("'Приказ о назначении сотрудника' идентификатор:")) &
         (df0['Ошибка'].str.contains("На дату кадрового приказа не найден действующий сотрудник."))]
df1.to_excel('join_' + '.xlsx', index=False)
print('\n=ИЛИ=\n', df1['Ошибка'])
 #На дату кадрового приказа не найден действующий сотрудник.
print('\n=И=\n',df1_['Ошибка'])
