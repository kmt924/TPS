import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText 
import re
from tkinter.messagebox import showinfo
import os


root = Tk()
root.title('Формирование файлов в Техподдержку ЕИСУКС')
root.geometry('850x600')



frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame.pack(anchor=NW, fill=X, padx=5, pady=5)

##############################################################################
    #  Checkbutton   "Выбор Вида Ошибок(Вид Файла)"


xls = pd.ExcelFile('виды_ошибок.xlsx')
sheets = xls.sheet_names
list_cb = []
print('44444', list_cb)
#VidF = sheets[1]
for j in range(len(sheets)):
    #if list_cb[j] != 'ошибки':
    list_cb.append(IntVar())
    print(list_cb[j])
def print_list_cb():
#    text.delete(1.0, END)
#    text.insert(1.0, list_cb)
    list_of_cb_values = []
    for i in range(len(sheets)):
        if sheets[i] != 'ошибки':
            list_of_cb_values.append(list_cb[i].get())
            print('в цикле: ', list_cb[i].get())
            if list_cb[i].get() == 1:
                print('в цикле фиксация Листа: ', list_cb[i].get())
                text.delete(1.0, END)
                text.insert(1.50, sheets[i])
                #VidF = sheets[i]
                #print('VidF', VidF)
            else:
                print('в цикле else фиксация Листа: ', list_cb[i].get())
                gg=0
                #text.delete(1.0, END)
                #text.insert(1.0, 'Выберите')
    #print(sheets[i], ff)
    #print(list_of_cb_values)
    #print(list_cb[i])        
for i in range(len(sheets)):
    if sheets[i] != 'ошибки':
        cb = Checkbutton(root, height=1, variable=list_cb[i], text=sheets[i],
                         command=print_list_cb)
        cb.pack(anchor=W, padx=5, pady=5)
#print('VidF_', VidF)
#btn = Button(root, text='print', command=print_list_cb)
#btn.pack(anchor=NW, padx=15, pady=1)

#############################################################################

text = ScrolledText(width=100, height=30, wrap="word") #  вертикальная прокрутка тхт-окна
#text.insert(7.0, '\n\nДля запуска программы нажмите СТАРТ.\n\n \n  \
#                      Программа работает с Отчётом =ЦА_ТО_Сведения_загрузки_данных=.\n\
#                      Отчёт должен находиться в одной папке с приложением.')
text.pack()


#print(len(sheets), sheets)
#for n in range(len(sheets)):
#    if sheets[n] != 'ошибки':
#        text.insert(1.0, sheets[n])
#        text.insert(1.0, '\n')
        
def clicked():
    print(os.getcwd())
    TEXTO_ = ''
    ii = 0
    VidF = text.get("1.0", "end")
    print('VidF_', VidF)
    #if VidF == '':                          # не работает код, если VidF пусто
    print('kkkkkkkkk', len(sheets), sheets)
    for n in range(len(sheets)):
        if sheets[n] != VidF:
            text.delete(1.0, END)
            text.insert(1.0, 'ОШИБКА!\nНЕ ВЫБРАН ТИП ДОКУМЕНТА!\n\
Выберите ТИП ДОКУМЕНТА для обработки в ЧекБоксе.')
    
    VidF =str(VidF.strip())
    #text.delete(1.0, END)
    print(VidF, enabled.get())
    
    df0 = pd.read_excel('ЦА_ТО_Сведения_загрузки_данных.xlsx', sheet_name='TDSheet')
    df1 = pd.read_excel('виды_ошибок.xlsx', sheet_name=VidF, usecols= 'B')
    # create an Empty DataFrame object
    column = df0.columns
    dff = pd.DataFrame(columns = column)
    print(df1.columns, len(df1))
    print(df0.columns, len(df0))
    
############################################
    
    for name in range(len(df1)):
        TEXTO = df1['text'].loc[df1.index[name]] # Выбор текста для поика
        #print('TEXTO  =  ', TEXTO)
        my_regex = r"\b(?=\w)" + re.escape(TEXTO) + r"\b(?!\w)" # регулярное выражение (формат для поиска)
        df0['Ошибка'].fillna('jjj', inplace=True) # замена NAN а текст 'jjj'
    
        for i in range(len(df0)):               # Просмотр всех строк фрейма df0
            df0_ = df0['Ошибка'].loc[df0.index[i]]
        #print(df0_)
            row = df0.iloc[i]
        #print(row[8])
            if re.search(my_regex, row[8]):     # Если ячейка 'Ошибка' совпадает с my_regex
                dff.loc[df0.index[i]] = df0.iloc[i]
                ii = ii +1
    #    print(row)
            TEXTO_1 = str(ii) + ' ' + TEXTO
        ii = 0
        TEXTO_ = TEXTO_ + '\n' +  TEXTO_1
    text.delete(1.0, END)      
    if enabled.get() ==1:
        dff = dff[dff['ВидФайла']. str.contains(VidF, na= False)]
        text.insert(END, 'Включен поиск по полю "Вид Файла"\n')  # Вывод поля "Вид файла"
    print(dff)
    dff = dff[['НомерОбласти','Учреждение','ИдентификаторУчреждения','ЦБ',
               'КодСВР','GUIDЗапроса','ВидФайла','ИдентификаторОбъекта','Ошибка']]
    
    dfff = dff[['ВидФайла','Ошибка']]
    
    text.insert(END, str(len(dfff)) + ' ' + VidF + '\n')     # Вывод "Тип Документа"
    text.insert(END, TEXTO_ +'\n')     # Вывод "Тип Документа" 
    text.insert(END, dfff)         # Вывод Отчета

    os.chdir('Отчеты')
    dff.to_excel(VidF + '.xlsx', index=False) # охранение в папку Отчеты
    os.chdir('..')
    
############################################
    d={}
    for name in range(len(df1)):
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
    
    # for i in str(name):
    #      d[name].to_excel(df1['text'].loc[df1.index[int(i)]] + '.xlsx', index=False)

    #print('Перс\n', df01)
##############################################################################
    #   Кнопки
          
Button(frame, text="СТАРТ", command=clicked).pack(side=LEFT)
Button(frame, text="Закрыть", command=root.destroy).pack(side=LEFT)

##############################################################################
    #  Checkbutton   "Включить для поиска поле "Вид Файла""
    
def checkbutton_changed():
    if enabled.get() == 1:
        xls = pd.ExcelFile('виды_ошибок.xlsx')
        sheets = xls.sheet_names
        print(len(sheets), sheets)
        #text.delete(1.0, END)
        #for n in range(len(sheets)):
        #    if sheets[n] != 'ошибки':
        #        text.insert(1.0, sheets[n])
        #        text.insert(1.0, '\n')
        #showinfo(title="Info", message="Включено")
    else:
        showinfo(title="Info", message="Отключен поиск по полю 'Вид Файла'.")
enabled = IntVar()
enabled_checkbutton = ttk.Checkbutton(frame, text='Включить для поиска поле "Вид Файла"',
                                      variable=enabled, command=checkbutton_changed)
enabled_checkbutton.pack(padx=16, pady=6)

######################

root.mainloop()
