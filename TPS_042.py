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
    #  Checkbutton 2   "Выбор Вида Ошибок(Вид Файла)"
frame6 = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame6.pack(anchor=NW, fill=X, padx=5, pady=5)
def selected(event):
    # получаем индексы выделенных элементов
    selected_indices = languages_listbox.curselection()
    # получаем сами выделенные элементы
    global selected_langs
    selected_langs = ",".join([languages_listbox.get(i) for i in selected_indices])
    msg = f"Вы выбрали: {selected_langs}"
    selection_label["text"] = msg
    text.delete(1.0, END)
    #text.insert(1.50, selected_langs)

VidF_df1 = pd.read_excel('виды_ошибок.xlsx', sheet_name='настройки_1')
print('Фрейм VidF_df1 (виды_ошибок.xlsx(лист "настройки_1"))\n', VidF_df1)
VidF1 = (VidF_df1.columns)
#print('Вид файла1:  \n', VidF1) 
list_cb1 = []  # список колонок (виды_ошибок.xlsx(лист "настройки_1")
for j in range(len(VidF_df1.columns)):
    if VidF1[j] != 'Вид файла:':
        list_cb1.append(VidF1[j])
print('PY_VAR0 (list_cb1)\n', list_cb1)
listbox_var = Variable(value=list_cb1)
print('listbox_var', listbox_var)

selection_label = ttk.Label(frame6, text="Выберите Вид Файла.")
selection_label.pack(anchor=NW, fill=X, padx=5, pady=5)
 
languages_listbox = Listbox(frame6, listvariable=listbox_var, bg = '#e9e9e9',
                            selectmode=SINGLE) # MULTIPLE) 
languages_listbox.pack(side=LEFT, fill=BOTH, expand=1) #anchor=NW, fill=X, padx=5, pady=5)
languages_listbox.bind("<<ListboxSelect>>", selected)
#languages_listbox.select_set(first=1)
#sd = languages_listbox.select_includes(3)
#print(sd)
scrollbar = ttk.Scrollbar(frame6, orient="vertical", command=languages_listbox.yview)
scrollbar.pack(side=RIGHT, fill=Y)
languages_listbox["yscrollcommand"]=scrollbar.set
##############################################################################
    #  Checkbutton   "Выбор Вида Ошибок(Вид Файла)"

#############################################################################
        
#  вертикальная прокрутка тхт-окна
text = ScrolledText(width=100, height=30, wrap="word") 
text.pack()

##############################################################################
    #  Button   "СТАРТ"
   
def clicked():
    def foo_2(arg):
        print(arg)
    if __name__ == '__main__':
        text.insert(1.0, 'ОШИБКА!\nНЕ ВЫБРАН ТИП ДОКУМЕНТА!\n\
Выберите ТИП ДОКУМЕНТА для обработки в ЧекБоксе.')    
        if selected_langs !='':
            print('+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\n',
                  selected_langs)
            VidF_in = selected_langs
        else:
            text.delete(1.0, END)
            text.insert(1.0, 'ОШИБКА!\nНЕ ВЫБРАН ТИП ДОКУМЕНТА!\n\
            Выберите ТИП ДОКУМЕНТА для обработки в ЧекБоксе.')      
            
    print(os.getcwd())
    #print('Вид файла:', VidF)
    #TEXTO_ = ''
    ii = 0
    VidF_in =str(VidF_in.strip())
    #text.delete(1.0, END)
    print('"Вид Файла" выбран', VidF_in, enabled.get())
    
    df0 = pd.read_excel('ЦА_ТО_Сведения_загрузки_данных.xlsx',
                        sheet_name='TDSheet')
    df1 = pd.read_excel('виды_ошибок.xlsx', sheet_name='настройки_1')
    # create an Empty DataFrame object
    column = df0.columns
    dff = pd.DataFrame(columns = column)
    ###########################################
    print('\ndf0.drop_duplicates()', len(df0))
    df0.drop_duplicates()
    print('\ndf0.drop_duplicates()', len(df0))
############################################
    print('Вид Файла и Поисковые Фразы. \n', VidF_in, '\n', df1[VidF_in])   
    text.delete(1.0, END)
    
############################################
    d=[] # пустой список для конкатеции
    d2=[]
    r=0
    result = pd.DataFrame(columns = column)
    #result_2 = pd.DataFrame
    for name in range(len(df1[VidF_in])):
        df1= df1. fillna('none')  ## замена NAN на текст '---'
        TEXTO = df1[VidF_in].loc[df1.index[name]] # Выбор текста для поика
        if TEXTO != '---':
            print('\n',name, 'TEXTO: ', df1[VidF_in].loc[df1.index[name]])
            df_1 = (df0[df0['Ошибка']. str.contains(
                    df1[VidF_in].loc[df1.index[name]], na= False)])
            #print('\nФрейм df_1',  len(df_1), '\n', df_1['Ошибка'])
            df_2 = df_1[['ВидФайла']]
            df_2.insert(1, "Поисковая фраза", TEXTO)
            if len(df_1) != 0:
                d.append(df_1)      ##  формирование списока для конкатеции
                d2.append(df_2)
            if TEXTO != 'none':
                text.insert(END, str(len(df_1)) + ' ' + TEXTO +'\n')     # Вывод "Тип Документа" 
    #print('\nСписок d\n', d)
    result = pd.concat(d) #, ignore_index=True)  ##  конкатеция
    result_2 = pd.concat(d2, ignore_index=True)  ##  конкатеция
    #print('result \n', result)

    if enabled.get() == 1:
        #result = result[result['ВидФайла']. str.contains(VidF_in, na= False)]
        result= result. fillna('none')
        result = result[(result['ВидФайла'].str.contains(VidF_in)) |
                            (result['ВидФайла'].str.contains('none'))]        
        #result_2 = result_2[result_2['ВидФайла']. str.contains(VidF_in, na= False)]
        result_2= result_2. fillna('none')
        result_2 = result_2[(result_2['ВидФайла'].str.contains(VidF_in)) |
                            (result_2['ВидФайла'].str.contains('none'))]
        text.insert('1.0', 'Включен поиск по полю "Вид Файла"\n')  # Вывод поля "Вид файла"
        #print(df_1)

##############################################################################
        # удаление 
    #print('\ndf0.drop_duplicates()', len(df0))
    #df0.drop_duplicates()
    print('\ndf0.drop_duplicates()', len(df0))
    print(result.index.tolist())
    result_ = df0.drop(result.index.tolist())
    print('\nresult_  ', len(result_))
    
    result = result[['НомерОбласти','Учреждение', 'ИдентификаторУчреждения','ЦБ','КодСВР',
                     'GUIDЗапроса','ВидФайла','ИдентификаторОбъекта', 'Ошибка']]    
    text.insert('1.0', str(len(result)) + ' ' + VidF_in + '\n')     # Вывод "Тип Документа"
    text.insert(END, result_2)         # Вывод Отчета
    print('\n result \n', result)
    result.to_excel('Свод_ошибок_по_учреждениям_' + VidF_in + '.xlsx', index=False)

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

    #else:
    #    showinfo(title="Info", message="Отключен поиск по полю 'Вид Файла'.")
enabled = IntVar()
enabled_checkbutton = ttk.Checkbutton(frame, text='Включить для поиска поле "Вид Файла"',
                                      variable=enabled, command=checkbutton_changed)
enabled_checkbutton.pack(padx=16, pady=6)

######################

root.mainloop()
