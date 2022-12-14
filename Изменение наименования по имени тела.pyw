#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     07.12.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "Изменение наименования по имени тела"
ver = "v0.4.0.0"

#------------------------------Настройки!---------------------------------------
recursive = False # рекурсивное (дет. внутри позсборок) переименование (True - да, False - нет)
Local_detail = True # обрабатывать локальные детали (True - да, False - нет)
No_MK = [11008, 11082, 11093, 11128, 11257, 11242, 11251, 11259] # перечень признаков что это не МК (см. ksObjectUserObject3D)
                                                                 # 11006 - эскиз; 11008 - элемент выдавливания; 11082 - массив по концентрической сетке; 11093 - зеркальный массив;
                                                                 # 11128 - операция вращения; 11200 - смещенная плоскость; 11210 - касательная плоскость; 11257 - фаска;
                                                                 # 11242 - условное изображение резьбы; # 11251 - отверстие; 11259 - скругление;
delete_names_MK = ["S = \d+ мм"] # список удаляемых слов/строк из названия МК (см. re), пример: ["S = \d+ мм", "@/L = \d+ мм"]
rounding_mass = 1 # округление массы (знаков после ",")
#-------------------------------------------------------------------------------

def KompasAPI(): # подключение API компаса

    from win32com.client import Dispatch, gencache # библиотека API Windows
    import pythoncom # модуль для запуска без IDLE
    from sys import exit # для выхода из приложения без ошибки

    try: # попытаться подключиться к КОМПАСу

        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
        global iKompasDocument # значение делаем глобальным
        global iPropertyMng # значение делаем глобальным

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
        KompasConst2D = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа 2D документов
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения

        KompasAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch('Kompas.Application.5', None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch('Kompas.Application.7') # интерфейс приложения КОМПАС-3D.

        iPropertyMng = KompasAPI7.IPropertyMng(iApplication) # интерфейс Менеджера свойств

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось подключиться к КОМПАСу

        Message("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def Message(text = "Ошибка!", counter = 4): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    from threading import Thread # библиотека потоков
    import time # модуль времени

    def Resource_path(relative_path): # для сохранения картинки внутри exe файла

        import os # работа с файовой системой

        try: # попытаться определить путь к папке
            base_path = sys._MEIPASS # путь к временной папки PyInstaller

        except Exception: # если ошибка
            base_path = os.path.abspath(".") # абсолютный путь

        return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

    def Message_Thread(text, counter): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        import tkinter.messagebox as mb # окно с сообщением
        import tkinter as tk # модуль окон

        if counter == 0: # время до закрытия окна (если 0)
            counter = 1 # закрытие через 1 сек
        window_msg = tk.Tk() # создание окна
        try: # попытаться использовать значёк
            window_msg.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        except: # если ошибка
            pass # пропустить
        window_msg.attributes("-topmost",True) # окно поверх всех окон
        window_msg.withdraw() # скрываем окно "невидимое"
        time = counter * 1000 # время в милисекундах
        window_msg.after(time, window_msg.destroy) # закрытие окна через n милисекунд
        if mb.showinfo(title, text, parent = window_msg) == "": # информационное окно закрытое по времени
            pass # пропустить
        else: # если не закрыто по времени
            window_msg.destroy() # окно закрыто по кнопке
        window_msg.mainloop() # отображение окна

    msg_th = Thread(target = Message_Thread, args = (text, counter)) # запуск окна в отдельном потоке
    msg_th.start() # запуск потока

    msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

def Kompas_message(text): # сообщение в окне КОМПАСа если он открыт

    if iApplication.Visible == True: # если компас видимый
        iApplication.MessageBoxEx(text, 'Message:', 64) # сообщение в КОМПАСе
    else: # если компас невидимый
        Message(text) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

def Main_Assembly(): # переименование деталей из сборки и её подсборок

    import re # модуль регулярных выражений

    global score # значение делаем глобальным

    iKompasDocument3D = KompasAPI7.IKompasDocument3D(iKompasDocument) # интерфейс документов-моделей
    iPart7 = iKompasDocument3D.TopPart # интерфейс компонента 3D документа (сам документ)

    if iPart7.Detail: # если дет.

        MK = Check_MK(iPart7) # определение МК

        if MK: # если это МК

            iName = iPart7.Name # наименование дет.

            iFeature7 = KompasAPI7.IFeature7(iPart7) # интерфейс объекта Дерева построения
            iName_body = iFeature7.ResultBodies.Name # имя тела дет.

            for delete_name_MK in delete_names_MK: # проходим по всему списку удаляемых слов/строк
                iName_body = re.sub(delete_name_MK, "", iName_body) # меняем слово/строку на "" (пустое)
                iName_body = iName_body.strip() # удаляем пробелы по краям строки

            Сhange = Сhange_properties(iPart7) # изменить св-ва

            if iName != iName_body or Сhange: # если наименование дет. и имя тела дет. разные и срабоал тригер изменения св-в

                iPart7.Name = iName_body # имя компонент
                iPart7.Update() # применить наименование

                score += 1 # добавляем счёт обработаных дет.

    else: # если это СБ

        Collect_Sources(iPart7) # рекурсивный сбор дет. и СБ

        if score != 0: # если был переименована хоть одна дет.
            iKompasDocument3D.RebuildDocument() # перестроить СБ
            iKompasDocument3D.Save() # сохранить изменения

def Check_MK(iPart7): # определение МК

    iModelContainer = KompasAPI7.IModelContainer(iPart7) # интерфейс контейнера трехмерных объектов
    iObjects = iModelContainer.Objects(0) # трехмерные объекты, входящие в состав данного объекта (объекты дерева построения)

    MK = 0 # количество МК элементов

    for iObject in iObjects: # перебор всех объектов

        if iObject.Type == 11211: # если найдена вставка с библиотеки МК
            MK += 1 # считаем количество вставок

            if MK == 2: # если вставок уже 2
                MK = False # не считаем это МК
                break # прерываем цикл

        if iObject.Type in (No_MK): # если в дереве построения есть признак
            MK = False  # не считаем это МК
            break # прерываем цикл

    else: # если цикл не прервался

        if MK == 0: # если не найдены втавка
            MK = False  # не считаем это МК

        else: # найдена одна вставка
            MK = True # это МК

    return MK # возвращаем значение

def Сhange_properties(iPart7): # изменить св-ва (значение формата, значение примечаний, значение массы, интерфейс формата, интерфейс примечания)

    iPropertyKeeper = KompasAPI7.IPropertyKeeper(iPart7) # интерфейс получения/редактирования значения свойств

    iProperty_sheet_formats = iPropertyMng.GetProperty(iKompasDocument, "Форматы листов документа") # интерфейс свойства
    sheet_formats = iPropertyKeeper.GetPropertyValue(iProperty_sheet_formats, 0, True)[1] # получить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))

    iProperty_note = iPropertyMng.GetProperty(iKompasDocument, "Примечание") # интерфейс свойства
    note = iPropertyKeeper.GetPropertyValue(iProperty_note, 0, True)[1] # получить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))

    iProperty_mass = iPropertyMng.GetProperty(iKompasDocument, "Масса") # интерфейс свойства
    mass = iPropertyKeeper.GetPropertyValue(iProperty_mass, 0, True)[1] # получить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))

    if rounding_mass != 0: # если округляем не до целого
        mass = round(mass, rounding_mass) # округляем значения массы
    else: # округляем до целого
        mass = round(mass) # округляем значения массы

    mass = str(mass).replace(".", ",") # меняем "." на ","

    Сhange = False # тригер изменения

    if sheet_formats != "БЧ": # если не записанно "БЧ" в формате
        iPropertyKeeper.SetPropertyValue(iProperty_sheet_formats, "БЧ", True) # установить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
        Сhange = True # тригер изменения

    if note != mass: # если отличается масса в примечании
        iPropertyKeeper.SetPropertyValue(iProperty_note, mass, True) # установить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
        Сhange = True # тригер изменения

    return Сhange

def Collect_Sources(iPart7): # рекурсивное переименование деталей дет. из подсборок

    def Rename_detail(iPart7): # переименование дет.

        import re # модуль регулярных выражений

        global score # значение делаем глобальным

        MK = Check_MK(iPart7) # определение МК

        if MK: # если это МК

            iSourcePart7Params = KompasAPI7.ISourcePart7Params(iPart7) # интерфейс параметров компонента в источнике
            iSourceName = iSourcePart7Params.SourceName # наименование дет.

            iFeature7 = KompasAPI7.IFeature7(iPart7) # интерфейс объекта Дерева построения
            iName_body = iFeature7.ResultBodies.Name # имя тела дет.

            for delete_name_MK in delete_names_MK: # проходим по всему списку удаляемых слов/строк
                iName_body = re.sub(delete_name_MK, "", iName_body) # меняем слово/строку на "" (пустое)
                iName_body = iName_body.strip() # удаляем пробелы по краям строки

            Сhange = Сhange_properties(iPart7) # изменить св-ва (интерфейс получения/редактирования значения свойств, значение формата, значение примечаний, значение массы, интерфейс формата, интерфейс примечания)

            if iSourceName != iName_body or Сhange: # если наименование дет. и имя тела дет. разные и срабоал тригер изменения св-в

                iSourcePart7Params.SourceName = iName_body # записываем имя компонента в источнике

                score += 1 # добавляем счёт обработаных дет.

    iPartsEx = iPart7.PartsEx(1) # список компонентов, включённыхв расчёт (0 - все компоненты (включая копии из операций копирования); 1 - первые экземпляры вставок компонентов (ksPart7CollectionTypeEnum))

    for iPart7 in iPartsEx: # проверяем каждый элемент из вставленных в СБ

        if iPart7.Detail: # если это дет.

            if iPart7.Standard == False: # если это не стандартная дет.

                if iPart7.IsLayoutGeometry == False: # если это не компоновочная геометрия

                    if iPart7.IsBillet == False: # если это не вставка заготовки дет.

                        if iPart7.IsLocal == False or Local_detail: # если это не локальная дет. или обрабатывать локальные детали включена
                            Rename_detail(iPart7) # переименование дет.

        else: # если это СБ
            if recursive: # если включён рекурсивное переименоване
                Collect_Sources(iPart7) # рекурсивный перебор

#-------------------------------------------------------------------------------
score = 0 # число обработаных дет.

KompasAPI() # подключение API компаса

Main_Assembly() # переименование деталей из сборки и её подсборок

if score == 0: # если нет переименованых дет.
    Kompas_message("Нет изменённых дет.") # сообщение окне КОМПАСа если он открыт
else: # если есть переименованые дет.
    Kompas_message("Изменено дет.: " + str(score)) # сообщение окне КОМПАСа если он открыт