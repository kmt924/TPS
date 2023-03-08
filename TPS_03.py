import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText 
import re
from tkinter.messagebox import showinfo
import os
import TPS_otmena
print(os.path)
#os.chdir(os.path) 
root = Tk()
root.title('Формирование файлов в Техподдержку ЕИСУКС')
root.geometry('850x650')



frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame.pack(anchor=NW, fill=X, padx=5, pady=5)

##############################################################################
    #  Checkbutton   "Выбор Вида Ошибок(Вид Файла)"
frame1 = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame1.pack(anchor=NW, fill=X, padx=5, pady=5)
frame2 = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame2.pack(anchor=NW, fill=X, padx=5, pady=5)
frame3 = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame3.pack(anchor=NW, fill=X, padx=5, pady=5)
#df1 = pd.read_excel('виды_ошибок.xlsx', sheet_name='настройки_1', usecols= 'B')
#VidF_df = pd.read_excel('виды_ошибок.xlsx', sheet_name='настройки_1', header=1)
VidF_df = pd.read_excel('виды_ошибок.xlsx', sheet_name='настройки_1')
print('Фрейм VidF_df \n', VidF_df)
global VidF 
VidF = VidF_df.columns
print('Вид файла:  \n', VidF)
list_cb = []
for j in range(len(VidF_df.columns)):
    list_cb.append(IntVar())
    print('IntVar', list_cb[j])
def print_list_cb():
    list_of_cb_values = []
    for i in range(len(VidF)):
        if VidF[i] != 'Вид файла:':
            list_of_cb_values.append(list_cb[i].get())
            # print('в цикле: ', list_cb[i].get())
            if list_cb[i].get() == 1:
                print('в цикле фиксация Листа: ', list_cb[i].get())
                text.delete(1.0, END)
                text.insert(1.50, VidF[i])
            else:
                # print('в цикле else фиксация Листа: ', list_cb[i].get())
                gg=0
       
for i in range(len(VidF)):
    if VidF[i] != 'Вид файла:':
        if i <=5:
            cb = Checkbutton(frame1, height=1, variable=list_cb[i],
                             text=VidF[i], command=print_list_cb)
            cb.pack(side=LEFT)
        if 5 < i <=10:
            cb = Checkbutton(frame2, height=1, variable=list_cb[i],
                             text=VidF[i], command=print_list_cb)
            cb.pack(side=LEFT)
        if 10 < i <=15:
            cb = Checkbutton(frame3, height=1, variable=list_cb[i],
                             text=VidF[i], command=print_list_cb)
            cb.pack(side=LEFT)
#############################################################################
        
#  вертикальная прокрутка тхт-окна
text = ScrolledText(width=100, height=30, wrap="word") 
               
text.pack()


  
def clicked():
    print(os.getcwd())
    #print('Вид файла:', VidF)
    TEXTO_ = ''
    ii = 0
    VidF_in = text.get("1.0", "end")
    print('VidF_in', VidF_in)
    #if VidF == '':                          # не работает код, если VidF пусто
    #    print('kkkkkkkkk', len(VidF), VidF)

    
    VidF_in =str(VidF_in.strip())
    #text.delete(1.0, END)
    print('"Вид Файла" выбран', VidF_in, enabled.get())
    
    df0 = pd.read_excel('ЦА_ТО_Сведения_загрузки_данных.xlsx',
                        sheet_name='TDSheet')
    df1 = pd.read_excel('виды_ошибок.xlsx', sheet_name='настройки_1')
    # create an Empty DataFrame object
    column = df0.columns
    dff = pd.DataFrame(columns = column)
    print(df1.columns, len(df1))
    print(df0.columns, len(df0))



    for n in range(len(VidF_in)):
        if VidF_in[n] != VidF_in:
            text.delete(1.0, END)
            text.insert(1.0, 'ОШИБКА!\nНЕ ВЫБРАН ТИП ДОКУМЕНТА!\n\
Выберите ТИП ДОКУМЕНТА для обработки в ЧекБоксе.')

############################################
    print('Вид Файла и Поисковые Фразы. \n', VidF_in, '\n', df1[VidF_in])
    for name in range(len(df1[VidF_in])): 
        df1[VidF_in].fillna('---', inplace=True) # замена NAN а текст '---'
        TEXTO = df1[VidF_in].loc[df1.index[name]] # Выбор текста для поиска
        if TEXTO != '---':
            print('TEXTO  =  ', TEXTO)
            my_regex = r"\b(?=\w)" + re.escape(TEXTO) + r"\b(?!\w)" # регулярное выражение (формат для поиска)
            df0['Ошибка'].fillna('---', inplace=True) # замена NAN на '---'
    
            for i in range(len(df0)):               # Просмотр всех строк фрейма df0
                #df0_ = df0['Ошибка'].loc[df0.index[i]]
                row = df0.iloc[i]                   # ячейка 'Ошибка'
                if re.search(my_regex, row[8]):     # Если ячейка 'Ошибка' содержит my_regex
                    dff.loc[df0.index[i]] = df0.iloc[i]  # Добавляем в фрейм dff строку из df0
                    dfff = dff[['ВидФайла']]
                    ii = ii + 1
                #dfff.insert(1, "Поисковая фраза", TEXTO, True)    
                TEXTO_1 = str(ii) + ' ' + TEXTO
            ii = 0
            TEXTO_ = TEXTO_ + '\n' +  TEXTO_1
    text.delete(1.0, END)      
    if enabled.get() ==1:
        dff = dff[dff['ВидФайла']. str.contains(VidF_in, na= False)]
        text.insert(END, 'Включен поиск по полю "Вид Файла"\n')  # Вывод поля "Вид файла"
    print('22222222222', len(dff), '\n', dff)
    dff = dff[['НомерОбласти','Учреждение','ИдентификаторУчреждения','ЦБ',
               'КодСВР','GUIDЗапроса','ВидФайла','ИдентификаторОбъекта','Ошибка']]
    
    dfff = dff[['ВидФайла','Ошибка']]
    #dfff.insert(1, "Поисковая фраза", TEXTO_1, True)
    text.insert(END, str(len(dfff)) + ' ' + VidF_in + '\n')     # Вывод "Тип Документа"
    text.insert(END, TEXTO_ +'\n')     # Вывод "Тип Документа" 
    text.insert(END, dfff)         # Вывод Отчета

    os.chdir('Отчеты')
    if len(dff) != 0:
        dff.to_excel(VidF_in + '.xlsx', index=False) # Сохранение в папку Отчеты
    os.chdir('..')
    
############################################
    d={}
    for name in range(len(df1[VidF_in])):
        df1[VidF_in].fillna('---', inplace=True) # замена NAN а текст '---'
        TEXTO = df1[VidF_in].loc[df1.index[name]] # Выбор текста для поика
        if TEXTO != '---':
            print(df1[VidF_in].loc[df1.index[name]])
        #df_1 = df0[df0['Ошибка']. str.contains(
        #    df1[VidF_in].loc[df1.index[name]], na= False)]
        #print('Персональный\n', df_1['Ошибка'])
            d[name] = df0[df0['Ошибка'].str.contains(
                df1[VidF_in].loc[df1.index[name]], na= False)]
            d[name] = d[name][['НомерОбласти','Учреждение',
                                   'ИдентификаторУчреждения','ЦБ','КодСВР',
                                   'GUIDЗапроса','ВидФайла','ИдентификаторОбъекта',
                                   'Ошибка']]
            print(name, len(d[name]), '  - Ожидаем_ ViewOrder_1')

# operate on DataFrame 'df' for company 'name'
    for name, df in d.items():
        print('555555555', df)
    
    ## for i in str(name):
    ##      d[name].to_excel(df1['text'].loc[df1.index[int(i)]] + '.xlsx', index=False)

    #print('Перс\n', df01)

##############################################################################

def otmena():

    text.delete(1.0, END)
    #label['text'] = ''
 
    s = TPS_otmena.otmena()
    text.insert(1.0, s)

##############################################################################
    #   clickedCheck   == ОСТАТКИ ==
    
def clickedCheck():
    text.delete(1.0, END)

    StrInReports = 0
    df1 = pd.read_excel('ЦА_ТО_Сведения_загрузки_данных.xlsx', sheet_name='TDSheet')
    df1 = df1[['НомерОбласти','Учреждение', 'ИдентификаторУчреждения','ЦБ','КодСВР',
                 'GUIDЗапроса','ВидФайла','ИдентификаторОбъекта', 'Ошибка']]
    #df1 = df1[['GUIDЗапроса',]]
    DIRE = os.listdir('Отчеты')
    print(len(df1))
    text.insert(END, 'Всего строк в исходном отчете: ')  # Вывод Строк в исходном отчете
    text.insert(END, len(df1)) 
    print(DIRE)
    column = df1.columns
    df_Rez = pd.DataFrame(columns = column)
    df_Rez_ = pd.DataFrame(columns = column)
    for file in range(len(DIRE)) :
        if re.search('.xlsx', DIRE[file]):
            if DIRE[file] != 'Остатки.xlsx':
                print(len(DIRE[file]), DIRE[file])
                text.insert(END, '\n\n')
                text.insert(END, DIRE[file]) # Вывод Имя Файла Отчета
                os.chdir('Отчеты')
                df2 = pd.read_excel(DIRE[file], sheet_name='Sheet1')
                df2 = df2[['НомерОбласти','Учреждение', 'ИдентификаторУчреждения','ЦБ','КодСВР',
                                     'GUIDЗапроса','ВидФайла','ИдентификаторОбъекта', 'Ошибка']]

                ee = ['НомерОбласти','Учреждение', 'ИдентификаторУчреждения','ЦБ','КодСВР',
                                     'GUIDЗапроса','ВидФайла','ИдентификаторОбъекта', 'Ошибка']
                StrInReports = StrInReports + len(df2)
                text.insert(END, '\nСтрок в отчете: ')  # Вывод Строк в отчете
                text.insert(END, len(df2))            
                os.chdir('..')
                for line_ in range(len(df2)):
                    for line in range(len(df1)):
                        df1 = df1.fillna(0)
                        df2 = df2.fillna(0)
                        #print (df1.iloc[line])
                        #print (df1[ee[1]].loc[df1.index[2]])
                        #print (df2.iloc[line_])
                        lnn = 0
                        for ln in range(8):
                            if (df1[ee[ln]].loc[df1.index[line]]) == (df2[ee[ln]].loc[df2.index[line_]]):
                                lnn = lnn + 1
                                if lnn < 8:
                                    df_Rez.loc[df1.index[line]] = df1.iloc[line]  # Добавляем в фрейм df_Rez строку из df1
                            lnn = 0
                print(len(df_Rez))

    text.insert(END, '\n\nСумма строк в Сводах: ') 
    text.insert(END, StrInReports)
    #df_Rez = pd.concat([df1,df_Rez]).drop_duplicates(keep=False)
    print(len(df_Rez))
    text.insert(END, '\n\nРазница Исходного и Сводов: ') 
    text.insert(END, len(df_Rez))

    os.chdir('Отчеты')
    df_Rez.to_excel('Остатки.xlsx', index=False) # Сохранение в папку Отчеты
    os.chdir('..')
    
##############################################################################
    #   Кнопки
          
Button(frame, text="СТАРТ", command=clicked).pack(side=LEFT)
Button(frame, text="Отмены", command=otmena).pack(side=LEFT)
Button(frame, text="Остатки", command=clickedCheck).pack(side=LEFT)
Button(frame, text="Закрыть", command=root.destroy).pack(side=LEFT)

value_var = IntVar()
progressbar =  ttk.Progressbar(frame, orient="horizontal", variable=value_var)
progressbar.pack(side=RIGHT)
label = ttk.Label(frame, textvariable=value_var)
label.pack(side=RIGHT)

##############################################################################
    #  Checkbutton   "Включить для поиска поле "Вид Файла""
    
def checkbutton_changed():
    if enabled.get() == 1:
        xls = pd.ExcelFile('виды_ошибок.xlsx')
        sheets = xls.sheet_names
        print(len(sheets), sheets)

    else:
        showinfo(title="Info", message="Отключен поиск по полю 'Вид Файла'.")
enabled = IntVar()
enabled_checkbutton = ttk.Checkbutton(frame, text='Включить для поиска поле "Вид Файла"',
                                      variable=enabled, command=checkbutton_changed)
enabled_checkbutton.pack(padx=16, pady=6)

######################

root.mainloop()
