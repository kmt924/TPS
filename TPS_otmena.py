import pandas as pd
import os
import re

# cmd /c cd "C:\Users\mkurbashev\Desktop\Работа\TPS\" && C:\Users\mkurbashev\AppData\Local\Programs\Python\Python39\python.exe "$(FULL_CURRENT_PATH)" && pause && quit
# && pause && quit  Старт Python
# cmd /k cd "C:\Users\mkurbashev\Desktop\Работа\TPS\" && C:\Users\mkurbashev\AppData\Local\Programs\Python\Python39\python.exe "$(FULL_CURRENT_PATH)" 
def otmena():
    df0 = pd.read_excel('ЦА_ТО_Сведения_загрузки_данных.xlsx', sheet_name='TDSheet')

    column = df0.columns
    df0_ = pd.DataFrame(columns = column)
    df_Otmena = df0[df0['Ошибка']. str.contains
        ('Приказ об отмене не загружен. Не найдены сведения о приказе, который требуется отменить',
        na= False)] 
    # Приказы об отмене
    ##print('222 \n', df_Otmena)
    df_Otmena_ = df_Otmena['Ошибка'].tolist()
    ##print(len(df_Otmena_), '  - Приказы об отмене (всего)')
    print(len(df0))
    for i2 in range(len(df_Otmena_)): 
        poz1 = df_Otmena_[i2].index(' идентификатор: ')
        poz2 = df_Otmena_[i2].index(
        'Приказ об отмене не загружен. Не найдены сведения о приказе, который требуется отменить')
        ID = df_Otmena_[i2][poz1 + 16:poz2 - 3:]
        #print(ID)
        for i in range(len(df0)):               # Просмотр всех строк фрейма df0
            row = df0.iloc[i]                   # строка 'Ошибка'
            if ID == row[12]: 
                df0_.loc[df0.index[i]] = df0.iloc[i]  # Добавляем в фрейм dff строку из df0
                df0_.loc[df_Otmena.index[i2]] = df_Otmena.iloc[i2]
        #df0 = df0[df0.ИдентификаторОбъекта != ID]  # удаление (построчно) строк содержащих ID
    print(len(df0))
    ##print(len(df0_))
    ##print(df0_)
    df0_ = df0_[['НомерОбласти','Учреждение','ИдентификаторУчреждения','ЦБ',
               'КодСВР','GUIDЗапроса','ВидФайла','ИдентификаторОбъекта','Ошибка']]
    
    dfff = df0_[['ВидФайла','Ошибка',]]
    os.chdir('Отчеты')
    if len(df0) != 0:
        df0_.to_excel('Свод Отмененные' + '.xlsx', index=False) # Сохранение в папку Отчеты
        #df0.to_excel('Свод без Отмен' + '.xlsx', index=False)
    os.chdir('..')
    
##############################################################################
        # удаление 
    #print('\ndf0.drop_duplicates()', len(df0))
    #df0.drop_duplicates()
    print('\ndf0.drop_duplicates()', len(df0_))
    print(df0_.index.tolist())
    result_df0 = df0.drop(df0_.index.tolist())
    print('\nresult_  ', len(df0_))
    print(len(result_df0))
    return dfff
