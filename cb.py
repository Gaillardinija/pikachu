import requests # для запросов cо страницы сайта
from bs4 import BeautifulSoup # парсер для синтаксического разбора страницы HTML/XML
import shutil, os # для работы с ОС


# определяем базовые параметры
link1 = 'https://www.cbr.ru/currency_base/daily/'  # первая часть ссылки 
link2 = '?date_req=' + '%s'

Norwegian = []

import pandas as pd
datelist = pd.date_range(pd.datetime(2019, 1, 1), pd.datetime.today()).tolist()
for i in datelist:
    page = requests.get(link1 + link2 % (i.strftime("%d.%m.%Y")))
    soup = BeautifulSoup(page.text, 'html.parser') # читаем эту страницу
    l = soup.find_all('td')
    for j in range(len(l)):
        if l[j].text.rfind('Норвежских крон') != -1:
            Norwegian.append([i.strftime("%d.%m.%Y"), l[j+1].text])
            print('данные за ' + i.strftime("%d.%m.%Y") + ' добавлены')
    


link = 'https://www.cbr.ru/currency_base/daily/?date_req=19.01.2019' 
page = requests.get(link)
soup = BeautifulSoup(page.text, 'html.parser')
l = soup.find_all('td')
for l1 in l:
    a = l1.text
    
for i in range(len(l)):
    if l[i].text.rfind('Норвежских крон') != -1:
        Norwegian.append([l[i].text, l[i+1].text])
        
        
    for i in range(len(l)): # проверяем, содержит ли тэг со ссылкой наименование необходимого файла, по нахождению прерываем цикл
        :
            filenameGE_eur = soup.find_all('a')[i].get_text() # определяем имя скачиваемого файла в найденно тэге
            linkGE = link1 + soup.find_all('a')[i].get('href') + link4 # определяем ссылку на файл для скачивания
            r = requests.get(linkGE)  # запрашиваем файл для скачивания
            output = open(filenameGE_eur, 'wb')
            output.write(r.content)
            output.close()
            source_files = os.getcwd()
            shutil.move(source_files + '\\' + filenameGE_eur, destination_folder + '\\eur\\' + filenameGE_eur)
            print('файл', filenameGE_eur, 'скачен и сохранен в папке', destination_folder + '\\eur')
        if i == len(l)-1 and str(l[i]).rfind(filename_eur) == -1:
            print('искомое имя файла по ссылке', link, 'в отобранных тэгах не найдено')
            
            
# составляем список файлов в папке назначения
import os
files = os.listdir(destination_folder + '\\eur')
GE_eur = []

for f in files:
    GE_eur.append(f[9:].replace(filename_eur, ''))

now18 = {'выходной' : ['03', '04', '10', '11', '17', '18', '24', '25'], 'понедельник' : ['05', '12', '19', '26'], 'вторник' : ['06', '13', '20', '27'], 'среда' : ['07', '14', '21', '28'], 'четверг' : ['01', '08', '15', '22', '29'], 'пятница' : ['02', '09', '16', '23', '30'], 'суббота' : ['03', '10', '17', '24', '31'], 'воскресенье' : ['04', '11', '18', '25']}
RGEweek = {'выходной' : [], 'понедельник' : [], 'вторник' : [], 'среда' : [], 'четверг' : [], 'пятница' : [], 'суббота' : [], 'воскресенье' : []}    
days = 30
PlanGE = {}

import xlrd
for f in files:
    for GE in GE_eur:
        if GE in f:
            PlanGE[GE] = []
            file = xlrd.open_workbook(destination_folder + '\\eur\\' + f)
            sheet = file.sheet_by_index(0)
            for i in range(sheet.nrows):
                for j in range(sheet.ncols):
                    if sheet.cell_value(i, j) == 'Плановый объем производства, МВтЧ':
                        start = 0
                        for k in range(i+1, sheet.nrows):
                            start += sheet.cell_value(k, j)
                        PlanGE[GE].append(start)
                        j += 5

from matplotlib import pyplot as plt
import numpy as np
for GE in GE_eur:
    fig = plt.figure()
    ax = fig.add_subplot(111)
    ax.plot([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23], PlanGE[GE][0:24])
    ax.set_xticks([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23])
    ax.set_title(GE)
    ax.grid(True)
plt.savefig('filename1.png', dpi = 300)
plt.close(fig)
    
                     
                    
                        
            



# формируем ссылки на страницу отчета за весь месяц (из расчета 31 день) и скачиваем сибирь
linkpage = []
for j in link3:
    linkpage.append(link1 + link2 + j + link5)

for link in linkpage:
    page = requests.get(link) # обращаемся к странице по сформированной сслыке
    soup = BeautifulSoup(page.text, 'html.parser') # читаем эту страницу
    l = soup.find_all('a') # находим теги со ссылками и добавляем в перечень
    for i in l: # проверяем, содержит ли тэг со ссылкой наименование необходимого файла, по нахождению прерываем цикл
        if str(i).rfind(filename_sib) != -1:
            filenameGEsib = i.get_text() # определяем имя скачиваемого файла в найденно тэге
            linkGEsib = link1 + i.get('href') + link5 # определяем ссылку на файл для скачивания
            r = requests.get(linkGEsib)  # запрашиваем файл для скачивания
            output = open(filenameGEsib, 'wb')
            output.write(r.content)
            output.close()
            source_files = os.getcwd()
            shutil.move(source_files + '\\' + filenameGEsib, destination_folder + '\\sib\\' + filenameGEsib)
            print('файл', filenameGEsib, 'скачен и сохранен в папке', destination_folder + '\\sib')
        if i == len(l)-1 and str(i).rfind(filename_sib) == -1:
            print('искомое имя файла по ссылке', link, 'в отобранных тэгах не найдено')
            
            
# составляем список файлов в папке назначения
import os
files = os.listdir(destination_folder)

