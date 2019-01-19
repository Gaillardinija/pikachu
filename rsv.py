import shutil, os 

destination_folder = 'C:\\Users\\pikachu\\Documents\\Лена'


# задаем параметры
year = '2019'
month = '01'
days = ['01','02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']
hours = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]
gtp = {'PADSHIPY':[0, 0, 0 ,0, 0, 0], 'PENRGEC1':[0, 0, 0, 0, 0, 0], 'PENRGEC4':[0, 0, 0, 0, 0, 0], 'PSKVIMP3':[0, 0, 0, 0, 0, 0], 'PSKVIMP5':[0, 0, 0, 0, 0, 0], 'PSKVIMP7':[0, 0, 0, 0, 0, 0]}


class RSV:
    def __init__(self, month, days, gtp, filename):
        self.month = month
        self.days = days
        self.gtp = gtp
        self.filename = filename
        global destination_folder
        import xlrd
        global ws
        wb = xlrd.open_workbook(destination_folder + '\\' + self.filename)
        ws = wb.sheet_by_index(0)
        

    def coll(self):
        for i in range(ws.nrows):
            for j in range(ws.ncols):
                if ws.cell_value(i, j) == 'Направление сделки':
                    global coll_buy
                    coll_buy = j
                if ws.cell_value(i, j) == 'Итого' and ws.cell_value(i+1, j) == 'Объем, кВтч' and ws.cell_value(i+1, j+1) == 'Предварительные требования/ обязательства, руб.':
                    global coll_v_rsv
                    global coll_s_rsv
                    coll_v_rsv = j
                    coll_s_rsv = j + 1
                if ws.cell_value(i, j) == 'Объем  РД, кВтч':
                    global coll_rd
                    coll_rd = j
    
    def buy_rsv(self):
        for consumer in self.gtp:
            for i in range(ws.nrows):
                for j in range(ws.ncols):
                    if ws.cell_value(i, j) == consumer and ws.cell_value(i, coll_buy) == 'покупка':
                        self.gtp[consumer][0] += ws.cell_value(i, coll_v_rsv)
                        self.gtp[consumer][1] += ws.cell_value(i, coll_s_rsv)

    def sell_rsv(self):
        for consumer in self.gtp:
            for i in range(ws.nrows):
                for j in range(ws.ncols):
                    if ws.cell_value(i, j) == consumer and ws.cell_value(i, coll_buy) == 'продажа':
                        self.gtp[consumer][2] += ws.cell_value(i, coll_v_rsv)
                        self.gtp[consumer][3] += ws.cell_value(i, coll_s_rsv)
                        
    def rd(self):
        for consumer in self.gtp:
            for i in range(ws.nrows):
                for j in range(ws.ncols):
                    if ws.cell_value(i, j) == consumer and ws.cell_value(i, coll_rd) > 0:
                        self.gtp[consumer][4] += ws.cell_value(i, coll_rd)
                        self.gtp[consumer][5] += ws.cell_value(i, coll_rd)*0.67537
                        
    def avg_price(self):
        file_list = os.listdir(destination_folder)
        if self.filename == file_list[-1]:
            for value in self.gtp.values():
                print((value[1]-value[3]+value[5])/(value[0]-value[2]+value[4]))
            
    def sum_proverka(self):
        file_list = os.listdir(destination_folder)
        if self.filename == file_list[-1]:
            buy_v_rsv = 0
            buy_s_rsv = 0
            sell_v_rsv = 0
            sell_s_rsv = 0
            buy_v_rd = 0
            buy_s_rd = 0
            for value in self.gtp.values():
                buy_v_rsv += value[0]
                buy_s_rsv += value[1]
                sell_v_rsv += value[2]
                sell_s_rsv += value[3]
                buy_v_rd += value[4]
                buy_s_rd += value[5]
            print(buy_v_rsv, buy_s_rsv, sell_v_rsv, sell_s_rsv, buy_v_rd, buy_s_rd)
            
    def graf_price(self):
        global rsv_buy_graf
        global rsv_sell_graf
        rsv_buy_graf = []
        rsv_sell_graf = []
        for i in range(ws.nrows):
            for j in range(ws.ncols):
                if ws.cell_value(i, coll_buy) == 'покупка' and ws.cell_value(i, coll_rd) > 0 and ws.cell_value(i, coll_s_rsv) > 0:
                    rsv_buy_graf.append(ws.cell_value(i, 0), ws.cell_value(i, 3), ws.cell_value(i, coll_s_rsv)/ws.cell_value(i, coll_v_rsv))
                if ws.cell_value(i, coll_buy) == 'продажа':
                    rsv_sell_graf.append(ws.cell_value(i, 0), ws.cell_value(i, 3), ws.cell_value(i, coll_s_rsv)/ws.cell_value(i, coll_v_rsv))
        
             
        


                        
                    
for filename in os.listdir(destination_folder):
    if 'Реестр сделок по торгам' in filename:
        a = RSV(month, days, gtp, filename)                       
        a.coll()
        a.buy_rsv()
        a.sell_rsv()
        a.rd()
        a.avg_price()
        a.sum_proverka()
        a.graf_price()                    
                    
