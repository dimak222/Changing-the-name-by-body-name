#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     07.12.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "Изменение наименования по имени тела"
ver = "v0.2.0.0"

#------------------------------Настройки!---------------------------------------
recursive = False # рекурсивное (дет. внутри позсборок) переименование (True - да, False - нет)
name_body = ["Тело"] # перечисление наименования тел дет. которые не будут переименовываться, пример: ["Тело", "тело", "Кот"]
#-------------------------------------------------------------------------------

def KompasAPI(): # подключение API компаса

    from win32com.client import Dispatch, gencache # библиотека API Windows
    import pythoncom # модуль для запуска без IDLE
    from sys import exit # для выхода из приложения без ошибки

    try: # попытаться подключиться к КОМПАСу

        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
        global iKompasDocument # значение делаем глобальным

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
        KompasConst2D = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа 2D документов
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения

        KompasAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch('Kompas.Application.5', None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch('Kompas.Application.7') # интерфейс приложения КОМПАС-3D.

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

    global score # значение делаем глобальным

    iKompasDocument3D = KompasAPI7.IKompasDocument3D(iKompasDocument) # интерфейс документов-моделей
    iPart7 = iKompasDocument3D.TopPart # интерфейс компонента 3D документа (сам документ)

    if iPart7.Detail: # если дет.

        iName = iPart7.Name # наименование дет.

        iFeature7 = KompasAPI7.IFeature7(iPart7) # интерфейс объекта Дерева построения
        iName_body = iFeature7.ResultBodies.Name # имя тела дет.

        for n in name_body: # перебор всех наименований
            print(n)
            if iName_body.find(n) != -1: # если найдено совпадение
                MK = False # не переименовываем
                break # прекращаем цикл

        else: # если цикл закончился
            MK = True # переименовываем

        if iName != iName_body and MK: # если наименование дет. и имя тела дет. разные и нет неправильных названий
            iPart7.Name = iFeature7.ResultBodies.Name # имя компонент
            iPart7.Update() # применить наименование
            score += 1 # добавляем счёт обработаных дет.

    else: # если это СБ

        Collect_Sources(iPart7) # рекурсивный сбор дет. и СБ

        if score != 0: # если был переименована хоть одна дет.
            iKompasDocument3D.RebuildDocument() # перестроить СБ
            iKompasDocument3D.Save() # сохранить изменения

def Collect_Sources(iPart7): # рекурсивный сбор дет. и СБ

    global score # значение делаем глобальным

    iPartsEx = iPart7.PartsEx(1) # список компонентов, включённыхв расчёт (0 - все компоненты (включая копии из операций копирования); 1 - первые экземпляры вставок компонентов (ksPart7CollectionTypeEnum))
    for iPart7 in iPartsEx: # проверяем каждый элемент из вставленных в СБ

        if iPart7.Detail: # если это дет.

            if iPart7.Standard == False: # если это не стандартная дет.

                if iPart7.IsLayoutGeometry == False: # если это не компоновочная геометрия

                    if iPart7.IsLocal == False: # если это не локальная деталь

                        if iPart7.IsBillet == False: # если это не вставка заготовки детали

                            iSourcePart7Params = KompasAPI7.ISourcePart7Params(iPart7) # интерфейс параметров компонента в источнике
                            iSourceName = iSourcePart7Params.SourceName # наименование дет.

                            iFeature7 = KompasAPI7.IFeature7(iPart7) # интерфейс объекта Дерева построения
                            iName_body = iFeature7.ResultBodies.Name # имя тела дет.

                            for n in name_body: # перебор всех наименований
                                if iName_body.find(n) != -1: # если найдено совпадение
                                    MK = False # не переименовываем
                                    break # прекращаем цикл

                            else: # если цикл закончился
                                MK = True # переименовываем

                            if iSourceName != iName_body and MK: # если наименование дет. и имя тела дет. разные и нет неправильных названий

                                iSourcePart7Params.SourceName = iFeature7.ResultBodies.Name # записываем имя компонента в источнике

                                score += 1 # добавляем счёт обработаных дет.

        else: # если это СБ
            if recursive: # если включён рекурсивное переименоване
                Collect_Sources(iPart7) # рекурсивный перебор

#-------------------------------------------------------------------------------
score = 0 # число обработаных дет.

KompasAPI() # подключение API компаса

Main_Assembly() # переименование деталей из сборки и её подсборок

Kompas_message("Переименовано дет.: " + str(score)) # сообщение в окне КОМПАСа если он открыт