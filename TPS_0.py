import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText 
import re

root = Tk()
root.title('Формирование файлов в Техподдержку ЕИСУКС')
root.geometry('450x300')
frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame.pack(anchor=NW, fill=X, padx=5, pady=5)
text = ScrolledText(width=50, height=15, wrap="word") #  вертикальная прокрутка тхт-окна
text.insert(3.0, 'Для запуска программы нажмите СТАРТ.\n\n \n  \
    Программа работает с Отчётом =ЦА_ТО_Сведения_загрузки_данных=.\n\
    Отчёт должен находиться в одной папке с приложением.')
text.pack()


def clicked():
    
    text.delete(1.0, END)
    df0 = pd.read_excel('ЦА_ТО_Сведения_загрузки_данных.xlsx', sheet_name='TDSheet')
    df1 = pd.read_excel('виды_ошибок.xlsx', sheet_name='Отпуска', usecols= 'B')
    # create an Empty DataFrame object
    column = df0.columns
    dff = pd.DataFrame(columns = column)
    print(df1.columns, len(df1))
    print(df0.columns, len(df0))
    #df1_ = df1['text'].tolist()
    #print(df1_, len(df1), )
    dd = 'Период простоя должен быть в пределах одного месяца'
    TEXTO = df1['text'].loc[df1.index[0]]
        #print('TEXTO  =  ', TEXTO)
    my_regex = r"\b(?=\w)" + re.escape(TEXTO) + r"\b(?!\w)"
    df0['Ошибка'].fillna('jjj', inplace=True) # замена NAN а текст 'jjj'
    for i in range(len(df0)):    
        df0_ = df0['Ошибка'].loc[df0.index[i]]
        #print(df0_)
        row = df0.iloc[i]
        #print(row[8])
        if re.search(my_regex, row[8]):
            dff.loc[df0.index[i]] = df0.iloc[i]
    #    print(row)
    print(dff)

    
    d={}
    for name in range(len(df1)):
        #d[name] = pd.DataFrame()
        #print(name, d[name])
    
 ##       print(df1['text'].loc[df1.index[name]])

        df_1 = df0[df0['Ошибка']. str.contains(df1['text'].loc[df1.index[name]],
                                               na= False)]
 ##       print('Персональный\n', df_1['Ошибка'])

        d[name] = df0[df0['Ошибка'].str.contains(
            df1['text'].loc[df1.index[name]], na= False)]
        d[name] = d[name][['НомерОбласти','Учреждение',
                                   'ИдентификаторУчреждения','ЦБ','КодСВР',
                                   'GUIDЗапроса','ВидФайла','ИдентификаторОбъекта',
                                   'Ошибка']]
        print(name, len(d[name]), '  - Ожидаем_ ViewOrder_1')

# operate on DataFrame 'df' for company 'name'
    for name, df in d.items():
        print('555555555', df)
    
    for i in str(name):
          d[name].to_excel(df1['text'].loc[df1.index[int(i)]] + '.xlsx', index=False)

    #print('Перс\n', df01)

Button(frame, text="СТАРТ", command=clicked).pack(side=LEFT)
Button(frame, text="Закрыть", command=root.destroy).pack(side=LEFT)

root.mainloop()
