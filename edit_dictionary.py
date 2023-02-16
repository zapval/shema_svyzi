from tkinter import ttk
from tkinter import *
from docx import Document
from docx.opc.coreprops import CoreProperties
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from tkinter import messagebox
import sqlite3
import numpy as np
from PIL import Image, ImageDraw, ImageFont


class Dictionary:
    db_name = 'dictionary.db'

    def __init__(self, window):

        self.wind = window
        self.wind.title('Редактирование подразделений')

        # создание элементов для ввода значений
        frame = LabelFrame(self.wind, text = 'Введите новое Подразделение')
        frame.grid(row = 0, column = 0, columnspan = 3, pady = 20)
        Label(frame, text = '№ связи: ').grid(row = 1, column = 0)
        self.nomer = Entry(frame)
        self.nomer.grid(row = 1, column = 1)
        Label(frame, text = '№ подразделения: ').grid(row = 2, column = 0)
        self.nomerpod = Entry(frame)
        self.nomerpod.grid(row = 2, column = 1)
        Label(frame, text = 'Подразделение: ').grid(row = 3, column = 0)
        self.word = Entry(frame)
        self.word.grid(row = 3, column = 1)
        Label(frame, text = 'Позывной: ').grid(row = 4, column = 0)
        self.meaning = Entry(frame)
        self.meaning.grid(row = 4, column = 1)
        ttk.Button(frame, text = 'Сохранить', command = self.add_word).grid(row = 5, columnspan = 2, sticky = W + E)
        self.message = Label(text = '', fg = 'green')
        self.message.grid(row = 4, column = 0, columnspan = 3, sticky = W + E)
        # таблица значений
        columns = ("#1", "#2", "#3")
        self.tree = ttk.Treeview(height = 10, columns = columns)
        self.tree.grid(row = 4, column = 0, columnspan = 3)
        self.tree.heading('#0', text = '№ связи', anchor = CENTER)
        self.tree.heading('#1', text = '№ подразделения', anchor = CENTER)
        self.tree.heading('#2', text = 'Подразделение', anchor = CENTER)
        self.tree.heading('#3', text = 'Позывной', anchor = CENTER)


        def generate():
            #узнаем сколько всего у нас подразделений
            query = 'SELECT COUNT(*) FROM dictionary'
            db_rows = self.run_query(query)
            for row in db_rows:
                number_of_rows=row[0]
            #берем из базы упорядоченные позывные подразделений
            query = 'SELECT meaning FROM dictionary ORDER BY nomer ASC, word ASC'
            db_rows = self.run_query(query)
            group=[]
            i=0
            #присваиваем переменным названия подразделений
            for row in db_rows:
                group.append(row[0])
                i=i+1
            #узнаем из базы уникальные значения номеров связи
            query = 'SELECT DISTINCT nomer FROM dictionary ORDER BY nomer ASC'
            db_rows = self.run_query(query)
            nomer_sv=[]
            i=0
            #присваиваем переменным номера связи
            for row in db_rows:
                nomer_sv.append(row[0])
                i=i+1
            #узнаем сколько всего у нас уникальных номеров связи
            query = 'SELECT COUNT(DISTINCT nomer) FROM dictionary'
            db_rows = self.run_query(query)
            for row in db_rows:
                kol_vo_sv=row[0]
            #узнаем сколько количество подразделений в каждом уникальном номере связи
            query = 'SELECT COUNT(DISTINCT meaning) FROM dictionary GROUP BY nomer'
            db_rows = self.run_query(query)
            i=0
            kol_vo_p=[]
            #присваиваем переменным количество подразделений в каждой связи
            for row in db_rows:
                kol_vo_p.append(row[0])
                i=i+1

            #дополнительные значения для генерации таблицы
            x=100
            y=400
            a=100
            b=100
            sto=100
            dve=200
            tri=300
            rast_str=250
            z=40
            w=50
            q = 10
            dob=0
            i = 0
            e = 0
            long = sto+w+200
            nomer_svyzi = ''
            img = Image.new('RGB', (1200, 600), 'white')
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype("Roboto-Regular.ttf", 12, encoding='UTF-8')
            #рисуем таблицу
            while i < 2:
                draw.line(
                    xy = (
                        (a, b),
                        (a+x, b),
                        (a+x, b+y),
                        (a, b+y),
                        (a, b)
                    ), fill ='black', width = 3)
                draw.line(
                    xy = (
                        (a, b+x),
                        (a+x, b+x)
                    ), fill ='black', width = 3)
                draw.text(
                    (a+q,b+z),
                    nomer_svyzi,
                    font=font,
                    fill=('black')
                    )
                nomer_svyzi = 'Усл. № связи'
                a=a+x
                i += 1
            #рисуем линии
            i = 0
            while i < kol_vo_sv:
                draw.line(
                    xy = (
                        (sto+w, rast_str+dob),
                        (long, rast_str+dob)
                    ), fill ='black', width = 2)
                draw.text(
                    (sto+sto+w,rast_str+dob-20),
                    nomer_sv[i],
                    font=font,
                    fill=('black')
                    )
                while e < kol_vo_p[i]:
                    draw.line(
                        xy = (
                            (long, rast_str+dob),
                            (long+100, rast_str+dob)
                        ), fill ='black', width = 2)
                    long = long + 100
                    draw.line(
                        xy = (
                            (long, rast_str+dob-10),
                            (long, rast_str+dob+10)
                        ), fill ='black', width = 2)
                    e=e+1
                dob=dob+50
                i += 1
                e = 0
            #дальше рисуем таблицу
            a=b+tri
            i = 0
            while i < number_of_rows:
                draw.line(
                    xy = (
                        (a, b),
                        (a+x, b),
                        (a+x, b+y),
                        (a, b+y),
                        (a, b)
                    ), fill ='black', width = 3)
                draw.line(
                    xy = (
                        (a, b+x),
                        (a+x, b+x)
                    ), fill ='black', width = 3)
                draw.text(
                    (a+z,b+z),
                    str(group[i]),
                    font=font,
                    fill=('black')
                    )
                a=a+x
                i += 1

            img.save('shema_KSHM.png')
            img = Image.open('shema_KSHM.png')
            img.show()

            output = 'shema_KSHM.png'

            img.save(output)

        #генерация схемы связи
        def generatesh():
            #узнаем сколько всего у нас подразделений
            query = 'SELECT COUNT(*) FROM dictionary'
            db_rows = self.run_query(query)
            for row in db_rows:
                number_of_rows=row[0]
            #берем из базы упорядоченные названия подразделений
            query = 'SELECT word FROM dictionary ORDER BY nomer ASC, word ASC'
            db_rows = self.run_query(query)
            groups=[]
            i=0
            #присваиваем переменным названия подразделений
            for row in db_rows:
                groups.append(row[0])
                i=i+1
            #узнаем из базы уникальные значения номеров связи
            query = 'SELECT DISTINCT nomer FROM dictionary ORDER BY nomer ASC'
            db_rows = self.run_query(query)
            nomer_sv=[]
            i=0
            #присваиваем переменным номера связи
            for row in db_rows:
                nomer_sv.append(row[0])
                i=i+1
            #узнаем сколько всего у нас уникальных номеров связи
            query = 'SELECT COUNT(DISTINCT nomer) FROM dictionary'
            db_rows = self.run_query(query)
            for row in db_rows:
                kol_vo_sv=row[0]
            #узнаем сколько количество подразделений в каждом уникальном номере связи
            query = 'SELECT COUNT(DISTINCT meaning) FROM dictionary GROUP BY nomer'
            db_rows = self.run_query(query)
            i=0
            kol_vo_p=[]
            #присваиваем переменным количество подразделений в каждой связи
            for row in db_rows:
                kol_vo_p.append(row[0])
                i=i+1

            #дополнительные значения для генерации таблицы
            x=1000
            y=400
            a=1100
            b=200
            sto=100
            dve=200
            tri=300
            rast_str=250
            z=40
            w=50
            q = 10
            dob=0
            i = 0
            e = 0
            long = sto+w+200

            img = Image.new('RGB', (2000, 1500), 'white')
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype("Roboto-Regular.ttf", 12, encoding='UTF-8')

            #рисуем КНП батальона
            draw.line(
                xy = (
                    (a, b),
                    (a+400, b),
                    (a+400, b+600),
                    (a, b+600),
                    (a, b)
                ), fill ='black', width = 3)
            #флажок батальона
            draw.line(
                xy = (
                    (a+200, b),
                    (a+200, b-40),
                    (a+220, b-30),
                    (a+200, b-20)
                ), fill ='red', width = 3)
            #рисуем 2 КШМ
            while i < 2:
                draw.line(
                    xy = (
                        (a+33, b+200),
                        (a+70, b+175),
                        (a+183, b+175),
                        (a+183, b+225),
                        (a+70, b+225),
                        (a+33, b+200)
                    ), fill ='black', width = 3)
                draw.line(
                    xy = (
                        (a+60, b+210),
                        (a+71.5, b+190),
                        (a+83, b+210),
                        (a+60, b+210)
                    ), fill ='black', width = 2)
                draw.line(
                    xy = (
                        (a+150, b+210),
                        (a+161.5, b+190),
                        (a+173, b+210),
                        (a+150, b+210)
                    ), fill ='black', width = 2)
                k = 0
                while k < 2:
                    draw.line(
                        xy = (
                            (a+90+dob, b+222),
                            (a+101.5+dob, b+202),
                            (a+113+dob, b+222),
                            (a+90+dob, b+222)
                        ), fill ='black', width = 2)
                    draw.line(
                        xy = (
                            (a+90+dob, b+198),
                            (a+101.5+dob, b+178),
                            (a+113+dob, b+198),
                            (a+90+dob, b+198)
                        ), fill ='black', width = 2)
                    k += 1
                    dob = 30
                dob = 0
                a = a + 183
                i += 1

            #узнеаем колво повторяющихся подразделений
            a=1100
            b=200
            grpov = [group[1:10] for group in groups]
            dup = [x for x in grpov if grpov.count(x) > 1]
            spis_pod_sv = []
            i = 0
            sk = kol_vo_p[0]
            sn = 0
            while i < kol_vo_sv:
                spis_pod_sv.append(groups[sn:sk])
                sn = sk
                i = i + 1
                if i < kol_vo_sv:
                    sk = sk + kol_vo_p[i]
            #рисуем линии от КШМ батальона
            draw.line(
                xy = (
                    (a+71.5, b+210),
                    (a+71.5, y+60),
                    (x+130, y+60)
                ), fill ='black', width = 2)
            draw.ellipse((a+69, y+57.5, a+74, y+62.5), fill ='black')
            draw.line(
                xy = (
                    (a+71.5+183, b+210),
                    (a+71.5+183, y+60),
                    (x+130, y+60)
                ), fill ='black', width = 2)
            if kol_vo_sv > 1:
                draw.line(
                    xy = (
                        (a+101.5, b+222),
                        (a+101.5, y+360),
                        (x+130, y+360)
                    ), fill ='black', width = 2)
                draw.ellipse((a+99, y+357.5, a+104, y+362.5), fill ='black')
                draw.line(
                    xy = (
                        (a+101.5+183, b+222),
                        (a+101.5+183, y+360),
                        (x+130, y+360)
                    ), fill ='black', width = 2)
            #рисуем саму схему
            #алгоритм: берем название подразделение, сравниваем с дубликтами, если есть совпадения, то дорисовываем элемент. затем удаляем все совпадения из дубликата.
            i = 0
            while i < kol_vo_sv:
                dup = [x for x in spis_pod_sv[i] if spis_pod_sv[i].count(x) > 1]
                k = 0
                while k < kol_vo_p[i]:
                    draw.line(
                        xy = (
                            (x, y),
                            (x, y+200),
                            (x-100, y+200),
                            (x-100, y),
                            (x, y)
                        ), fill ='black', width = 3)
                    draw.text(
                        (x-60, y-15),
                        spis_pod_sv[i][k],
                        font=font,
                        fill=('black')
                        )
                    draw.line(
                        xy = (
                            (x+130, y+60),
                            (x-70, y+60)
                        ), fill ='black', width = 2)
                    if dup:
                        if spis_pod_sv[i][k] == dup[0]:
                            s = 0
                            p = spis_pod_sv[i].count(dup[0])
                            while s < p - 1:
                                draw.line(
                                    xy = (
                                        (x, y+10),
                                        (x+10, y+10),
                                        (x+10, y+210),
                                        (x-90, y+210),
                                        (x-90, y+200)
                                    ), fill ='black', width = 3)
                                x = x + 10
                                y = y + 10
                                s = s + 1
                            x = x - (10 * s) - 200
                            y = y - (10 * s)
                            k = k + p
                            del dup[0:p]
                        else:
                            x = x - 200
                            k = k + 1
                    else:
                        x = x - 200
                        k = k + 1
                x = 1000
                y = y + 300
                i = i + 1



            img.save('shema_svyzi.png')
            img = Image.open('shema_svyzi.png')
            img.show()

            output = 'shema_svyzi.png'

            img.save(output)

        # кнопки создания документа и удаления записей
        ttk.Button(text = 'Сгенерировать документ', command = generate).grid(row = 5, column = 0, columnspan = 1, sticky = W + E)
        ttk.Button(text = 'Сгенерировать схему', command = generatesh).grid(row = 5, column = 1, columnspan = 1, sticky = W + E)
        ttk.Button(text = 'Удалить', command = self.delete_word).grid(row = 5, column = 2, columnspan = 1, sticky = W + E)

        # заполнение таблицы
        self.get_words()


    # подключение и запрос к базе
    def run_query(self, query, parameters = ()):
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parameters)
            conn.commit()
        return result

    # заполнение таблицы значениями
    def get_words(self):
        records = self.tree.get_children()
        for element in records:
            self.tree.delete(element)
        query = 'SELECT * FROM dictionary ORDER BY nomer DESC, word DESC, nomerpod DESC'
        db_rows = self.run_query(query)
        for row in db_rows:
            self.tree.insert('', 0, text = row[1], values = (row[2], row[3], row[4]))

    # валидация ввода
    def validation(self):
        return len(self.word.get()) != 0 and len(self.meaning.get()) != 0
    # добавление нового подразделения
    def add_word(self):
        if self.validation():
            query = 'INSERT INTO dictionary VALUES(NULL, ?, ?, ?, ?)'
            parameters =  (self.nomer.get(), self.nomerpod.get(), self.word.get(), self.meaning.get())
            self.run_query(query, parameters)
            self.message['text'] = 'Подразделение {} добавлено'.format(self.word.get())
            self.nomer.delete(0, END)
            self.nomerpod.delete(0, END)
            self.word.delete(0, END)
            self.meaning.delete(0, END)
        else:
            self.message['text'] = 'введите слово и значение'
        self.get_words()
    # удаление подразделения
    def delete_word(self):
        self.message['text'] = ''
        try:
            self.tree.item(self.tree.selection())['text'][0]
        except IndexError as e:
            self.message['text'] = 'Выберите подразделение, которое нужно удалить'
            return
        self.message['text'] = ''
        meaning = self.tree.item(self.tree.selection())['values'][2]
        query = 'DELETE FROM dictionary WHERE meaning = ?'
        self.run_query(query, (meaning, ))
        self.message['text'] = 'Подразделение {} успешно удалено'.format(meaning)
        self.get_words()

window = Tk()
application = Dictionary(window)
window.mainloop()
