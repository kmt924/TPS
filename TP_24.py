import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText 


root = Tk()
root.title('Формирование файлов в Техподдержку ЕИСУКС')
root.geometry('450x300')

enabled = IntVar()
  
enabled_checkbutton = ttk.Checkbutton(text='Включить для поиска поле "Вид Файл"', variable=enabled)
enabled_checkbutton.pack(padx=6, pady=6, anchor=NW)
  
#enabled_label = ttk.Label(textvariable=enabled)
#enabled_label.pack(padx=6, pady=6, anchor=NW)


frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame.pack(anchor=NW, fill=X, padx=5, pady=5)
text = ScrolledText(width=50, height=15, wrap="word") #  вертикальная прокрутка тхт-окна
text.insert(3.0, 'Для запуска программы нажмите СТАРТ.\n\n \n  \
    Программа работает с Отчётом =ЦА_ТО_Сведения_загрузки_данных=.\n\
    Отчёт должен находиться в одной папке с приложением.')
text.pack()



xls = pd.ExcelFile('виды_ошибок.xlsx')
sheets = xls.sheet_names
print(sheets)


def clicked():
    #df1 = pd.read_excel('виды_ошибок.xlsx').sheet_names
    #print(df1)

    
    VidF = text.get("1.0", "end")
    VidF =str(VidF.strip())
    print(VidF, enabled)
    #text.delete(1.0, END)
    #label['text'] = ''    
    df0 = pd.read_excel('ЦА_ТО_Сведения_загрузки_данных.xlsx', sheet_name='TDSheet')
    df1 = pd.read_excel('виды_ошибок.xlsx', sheet_name=VidF, usecols= 'B')
    print(df1.columns, len(df1))
    df1_ = df1['ФразаДляПоиска'].tolist()
    print(df1_, len(df1))

    d = {}
    df =[]
    for name in range(len(df1_)):

        df_1 = df0[df0['Ошибка']. str.contains(df1['ФразаДляПоиска'].loc[df1.index[name]],
                                               na= False)]
        #print('Персональный\n', df_1['Ошибка'])

        d[name] = df0[df0['Ошибка'].str.contains(
            df1['ФразаДляПоиска'].loc[df1.index[name]], na= False)]
        #if enabled == 1:
        d[name] = d[name][d[name]['ВидФайла'].str.contains('Материальное стимулирование', na= False)]
        d[name] = d[name][['НомерОбласти','Учреждение',
                                   'ИдентификаторУчреждения','ЦБ','КодСВР',
                                   'GUIDЗапроса','ВидФайла','ИдентификаторОбъекта',
                                   'Ошибка']]
        print(name, len(d[name]), '  - Ожидаем_ ViewOrder_1\n', d)

        for i in str(name):
            #d[name].to_excel(df1['text'].loc[df1.index[int(i)]] + '.xlsx', index=False)
            d[name].to_excel(VidF +'_' + i + '.xlsx', index=False)


Button(frame, text="СТАРТ", command=clicked).pack(side=LEFT)
Button(frame, text="Закрыть", command=root.destroy).pack(side=LEFT)

root.mainloop()
