print('введите гп, регион, уровень напряжения, год')
gp = input()
reg = input()
volt = input()
year = str(input())

import mysql.connector
from mysql.connector import errorcode
ee_param = {'host': '127.0.0.1', 'user': 'elena', 'password': '1689170221', 'database': 'ee' }
conn = mysql.connector.connect(**ee_param)
cursor = conn.cursor()

_sql = 'SELECT  Code_Supplier_ee, N_Region from gp where Supplier_ee like %s and Region like %s'
cursor.execute(_sql, (gp, reg))
for i in cursor.fetchall():
    code_gp = i[0]
    code_reg = i[1]
print(code_gp)
print(code_reg)

_sql = 'SELECT ot670_do10, Period from sn where Supplier_ee like %s and Region like %s and Period like %s'
cursor.execute(_sql, (gp, reg, '%'+year))
sn = cursor.fetchall()
print(sn)

_sql = 'SELECT one_ee, two_ee, two_power_so, Period from grid_price where RealVoltage like %s and Region like %s and Period like %s'
cursor.execute(_sql, (volt, reg, '%'+year))
seti = cursor.fetchall()
print(seti)

_sql = 'SELECT inie, Period from inie_price where Supplier_ee like %s and Region like %s and Period like %s'
cursor.execute(_sql, (gp, reg, '%'+year))
inie = cursor.fetchall()
print(inie)

import numpy as np
import xlrd

datas = np.array([['январь', year + '01_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '0210_' + code_gp + '_' + 'PSAMARAE' + '_01' + year + '_gtp_1st_stage.xls'],
['февраль', year + '02_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '0310_' + code_gp + '_' + 'PSAMARAE' + '_02' + year + '_gtp_1st_stage.xls'],
['март', year + '03_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '0410_' + code_gp + '_' + 'PSAMARAE' + '_03' + year + '_gtp_1st_stage.xls'],
['апрель', year + '04_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '0510_' + code_gp + '_' + 'PSAMARAE' + '_04' + year + '_gtp_1st_stage.xls'],
['май', year + '05_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '0610_' + code_gp + '_' + 'PSAMARAE' + '_05' + year + '_gtp_1st_stage.xls'],
['июнь', year + '06_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '0710_' + code_gp + '_' + 'PSAMARAE' + '_06' + year + '_gtp_1st_stage.xls'],
['июль', year + '07_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '0810_' + code_gp + '_' + 'PSAMARAE' + '_07' + year + '_gtp_1st_stage.xls'],
['август', year + '08_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '0910_' + code_gp + '_' + 'PSAMARAE' + '_08' + year + '_gtp_1st_stage.xls'],
['сентябрь', year + '09_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '1010_' + code_gp + '_' + 'PSAMARAE' + '_09' + year + '_gtp_1st_stage.xls'],
['октябрь', year + '10_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '1110_' + code_gp + '_' + 'PSAMARAE' + '_10' + year + '_gtp_1st_stage.xls'],
['ноябрь', year + '11_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year + '1210_' + code_gp + '_' + 'PSAMARAE' + '_11' + year + '_gtp_1st_stage.xls'],
['декабрь', year + '12_' + code_gp + '_' + code_reg + '_calcfacthour.xls', year[0:2]+str(int(year[2:4])+1) + '1210_' + code_gp + '_' + 'PSAMARAE' + '_12' + year + '_gtp_1st_stage.xls']])
         

print(datas)

'''

count = 0
for key, value in months.items():
    wb = xlrd.open_workbook(destination_folder + '1.xlsx')
    if key == datas[count][0]:
        month = datas[count][0]
        ws = wb.sheet_by_name(month)
        val = [ws.cell_value(i, 0) for i in range(value*24)]
        val_ar = np.array(val)
        wb = xlrd.open_workbook(destination_folder + datas[count][1])
        ws = wb.sheet_by_index(0)
        for row in range(ws.nrows):
            if ws.cell_value(row, 0) == 'Дата':
                data_ATS = [[ws.cell_value(row1, 0), ws.cell_value(row1, 1)] for row1 in range(row+2, ws.nrows)]
                break
        power_ATS = sum([val[24*((int(data[0][0:2]))-1)+int(data[1])-1] for data in data_ATS])/len(data_ATS)
        print(power_ATS)
        power_SO = []
        for data in data_ATS:
            if len(hours_SO[month]) <= 2:
                start = 24*((int(data[0][0:2]))-1)+hours_SO[month][0]-1
                finish = 24*((int(data[0][0:2]))-1)+hours_SO[month][1]
                data_SO= [i for i in range(start, finish)]
                power_SO.append(max([val[data] for data in data_SO]))
            if len(hours_SO[month]) > 2:
                start1 = 24*((int(data[0][0:2]))-1)+hours_SO[month][0]-1
                finish1 = 24*((int(data[0][0:2]))-1)+hours_SO[month][1]
                start2 = 24*((int(data[0][0:2]))-1)+hours_SO[month][2]-1
                finish2 = 24*((int(data[0][0:2]))-1)+hours_SO[month][3]
                data_SO= [i for i in range(start1, finish1)] + [i for i in range(start2, finish2)]
                power_SO.append(max([val[data] for data in data_SO]))
        power_SO = sum(power_SO)/len(power_SO)
        print(power_SO)
        wb = xlrd.open_workbook(destination_folder + datas[count][2])
        ws = wb.sheet_by_index(0)
        for row in range(ws.nrows):
            for col in range(ws.ncols):
                if ws.cell_value(row, col) == 'Средневзвешенная нерегулируемая цена на мощность на оптовом рынке, руб/МВт':
                    costs_ATS = power_ATS*float(ws.cell_value(row, col+1).replace(',', '.'))/1000
                if 'Дифференцированная по часам расчетного периода нерегулируемая цена на электрическую энергию на оптовом рынке, определяемая по результатам конкурентного отбора ценовых заявок на сутки вперед и конкурентного отбора заявок для балансирования системы' in str(ws.cell_value(row, col)):
                    price_RSV = [float(ws.cell_value(row1, col).replace(',', '.'))/1000 for row1 in range(row+1, row+1+value*24)]
                    price_RSV_ar = np.array(price_RSV)
                    costs_ee = np.sum(val_ar*price_RSV_ar)
        if month in ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь']:
            price3 = costs_ee + costs_ATS + np.sum(sn_1[2]*(val_ar*price_RSV_ar)) + power_ATS*sn_1[2]*costs_ATS/power_ATS + price_seti_1[0]*np.sum(val_ar)+inie[month]*np.sum(val_ar)
            price[month + '3'] = [['Электроэнергия', 'кВтч', np.sum(val_ar), costs_ee/np.sum(val_ar),costs_ee],
                 ['Сбытовая надбавка ээ', 'кВтч', np.sum(val_ar), np.sum(sn_1[2]*(val_ar*price_RSV_ar))/np.sum(val_ar), np.sum(sn_1[2]*(val_ar*price_RSV_ar))],
                 ['Сбытовая надбавка мощность', 'кВт', power_ATS, sn_1[2]*costs_ATS/power_ATS, power_ATS*sn_1[2]*costs_ATS/power_ATS], 
                 ['Иные услуги', 'кВтч', np.sum(val_ar), inie[month], inie[month]*np.sum(val_ar)],
                 ['Услуга по передаче', 'кВтч', np.sum(val_ar), price_seti_1[0], price_seti_1[0]*np.sum(val_ar)],
                 ['Мощность ОРЭМ', 'кВт', power_ATS, costs_ATS/power_ATS, costs_ATS],
                 ['ИТОГО', 'кВтч', np.sum(val_ar), price3/np.sum(val_ar), price3]]
            price4 = costs_ee + costs_ATS + np.sum(sn_1[2]*(val_ar*price_RSV_ar)) + power_ATS*sn_1[2]*costs_ATS/power_ATS + price_seti_1[1]*np.sum(val_ar)+price_seti_2[2]*power_SO+inie[month]*np.sum(val_ar)
            price[month + '4'] = [['Электроэнергия', 'кВтч', np.sum(val_ar), costs_ee/np.sum(val_ar),costs_ee],
                 ['Сбытовая надбавка ээ', 'кВтч', np.sum(val_ar), np.sum(sn_1[2]*(val_ar*price_RSV_ar))/np.sum(val_ar), np.sum(sn_1[2]*(val_ar*price_RSV_ar))],
                 ['Сбытовая надбавка мощность', 'кВт', power_ATS, sn_1[2]*costs_ATS/power_ATS, power_ATS*sn_1[2]*costs_ATS/power_ATS], 
                 ['Иные услуги', 'кВтч', np.sum(val_ar), inie[month], inie[month]*np.sum(val_ar)],
                 ['Услуга по передаче э/э', 'кВтч', np.sum(val_ar), price_seti_1[1], price_seti_1[1]*np.sum(val_ar)],
                 ['Услуга по передаче мощность', 'кВт', power_SO, price_seti_1[2], price_seti_1[2]*power_SO],
                 ['Мощность ОРЭМ', 'кВт', power_ATS, costs_ATS/power_ATS, costs_ATS],
                 ['ИТОГО', 'кВтч', np.sum(val_ar), price4/np.sum(val_ar), price4]]       
        else:
            price3 = costs_ee + costs_ATS + sn_2[2]*np.sum(val_ar) + price_seti_2[0]*np.sum(val_ar)+inie[month]*np.sum(val_ar)
            price[month + '3'] = [['Электроэнергия', 'кВтч', np.sum(val_ar), costs_ee/np.sum(val_ar), costs_ee],
                 ['Сбытовая надбавка', 'кВтч', np.sum(val_ar), sn_2[2], sn_2[2]*np.sum(val_ar)],
                 ['Иные услуги', 'кВтч', np.sum(val_ar), inie[month], inie[month]*np.sum(val_ar)],
                 ['Услуга по передаче', 'кВтч', np.sum(val_ar), price_seti_2[0], price_seti_2[0]*np.sum(val_ar)],
                 ['Мощность ОРЭМ', 'кВт', power_ATS, costs_ATS/power_ATS, costs_ATS],
                 ['ИТОГО', 'кВтч', np.sum(val_ar), price3/np.sum(val_ar), price3]]
            price4 = costs_ee + costs_ATS + sn_2[2]*np.sum(val_ar) + price_seti_2[1]*np.sum(val_ar)+price_seti_2[2]*power_SO + inie[month]*np.sum(val_ar)
            price[month + '4'] = [['Электроэнергия', 'кВтч', np.sum(val_ar), costs_ee/np.sum(val_ar),costs_ee],
                 ['Сбытовая надбавка', 'кВтч', np.sum(val_ar), sn_2[2], sn_2[2]*np.sum(val_ar)],
                 ['Иные услуги', 'кВтч', np.sum(val_ar), inie[month], inie[month]*np.sum(val_ar)],
                 ['Услуга по передаче э/э', 'кВтч', np.sum(val_ar), price_seti_2[1], price_seti_2[1]*np.sum(val_ar)],
                 ['Услуга по передаче мощность', 'кВт', power_SO, price_seti_2[2], price_seti_2[2]*power_SO],
                 ['Мощность ОРЭМ', 'кВт', power_ATS, costs_ATS/power_ATS, costs_ATS],
                 ['ИТОГО', 'кВтч', np.sum(val_ar), price4/np.sum(val_ar), price4]]
    count += 1

        
import xlwt
wb =  xlwt.Workbook()
ws = wb.add_sheet('1')
i = 0
for key, value in price.items():
    ws.write(i, 0, key)  
    for val in value:
        ws.write(i, 1, val[0])
        ws.write(i, 2, val[1])  
        ws.write(i, 3, val[2])  
        ws.write(i, 4, val[3])  
        ws.write(i, 5, val[4])
        i += 1
wb.save(destination_folder + 'Свод.xls')  

'''