import time
import xlrd
import xlwt
import model
import optimized


class data:
    def __init__(self, month, day, length):
        self.month = month
        self.day = day
        self.length = length
    
    def setmonth(self, month):
        self.month = month
        self.length = 1  # 为了现实效果，把每个月的长度设置为了一天
    
    def setday(self, day):
        self.day = day
    
    def nextmonth(self):
        if self.month < 12:
            self.month += 1
        else:
            self.month = 1
        self.length = 1
    
    def nextday(self):
        if self.day < self.length:
            self.day += 1
        else:
            self.day = 1
            self.nextmonth()

# def test(city_willings):
#     for i in range(btime.month - 1, etime.month - 2):
#         if city_willings[i].statue != 0:
#             return 0
#         else:
#
def prewrite(row, col, num):
    a = [row, col, num]
    write_data.append(a)


if __name__ == '__main__':
    btime = data(0, 0, 0)
    ctime = data(0, 0, 0)
    etime = data(0, 0, 0)
    # 设置开始时间和结束时间
    btime.setday(optimized.btime.day)
    btime.setmonth(optimized.btime.month)
    etime.setday(optimized.etime.day)
    etime.setmonth(optimized.etime.month)
    # 设置预见的天数
    foresee = 4
    # 读取城市的数据
    citydata = xlrd.open_workbook("data.xlsx")
    city_sheets = citydata.sheet_by_name("city")
    citys_data = []  # 储存所有城市的数据，元素为city_data
    city_date = []  # 存了所有城市的十二个月的信息，元素为value（list）
    rows = city_sheets.nrows  # 获取行数
    for i in range(1, rows):
        row = city_sheets.row_values(i)
        value = []  # 存了一个月的信息，元素是str、num
        flag = 0
        for s in row:
            if s != '':
                value.append(s)
        city_date.append(value)
    citys_data = [city_date[i:i + 12] for i in range(0, len(city_date), 12)]
    write_data = []
    result_time = 0
    active = True
    citys = []
    station = model.peer("station", 10000000, 1000000, 0, 0, 0, 50, 0, 0, 10000000000)
    
    for s in citys_data.copy():
        city = model.peer(s[btime.month - 1][10], s[btime.month - 1][1], s[btime.month - 1][2], s[btime.month - 1][3],
                          s[btime.month - 1][4], s[btime.month - 1][5], s[btime.month - 1][6],
                          s[btime.month - 1][7], s[btime.month - 1][8], s[btime.month - 1][9])
        citys.append(city)
    
    # 从优化过的execl里读取willing
    willingdata = xlrd.open_workbook("optimized.xls")
    willing_sheets = willingdata.sheet_by_name("My Sheet")
    citys_willings = []
    city_willings = []
    for i in range(1, willing_sheets.nrows):
        row = willing_sheets.row_values(i)
        value = []
        flag = 0
        for s in row:
            if s != '':
                value.append(s)
        city_willings.append(value)
    citys_willings = [city_willings[i:i + (etime.month - btime.month)] for i in
                      range(0, len(city_willings), (etime.month - btime.month))]
    print(citys_willings[1][6])
    while False:
        result = []
        buy_willings = []
        sale_willings = []
        orderlist = []
        flag = 0
        for s in citys.copy():
            try:
                if CitysWillings[flag][btime.month - 1].statue == 1:  # 买入欲望
                    a = model.buy_willing(s, CitysWillings[flag][btime.month - 1].number)  # 发出购买欲望
                    buy_willings.append(a)
                    flag += 1
                elif CitysWillings[flag][btime.month - 1].statue == 0:
                    a = model.sale_willing(s, CitysWillings[flag][btime.month - 1].number, s.sale_price)
                    sale_willings.append(a)
                    flag += 1
                else:
                    pass
            except BaseException:
                pass
        while True:
            if buy_willings == [] or sale_willings == []:
                break
            if buy_willings[0].number > sale_willings[0].number:
                a = model.order(buy_willings[0].buyer, sale_willings[0].saler, sale_willings[0].price,
                                sale_willings[0].number)
                orderlist.append(a)
                buy_willings[0].reduce(sale_willings[0].number)
                sale_willings.remove(sale_willings[0])
            elif buy_willings[0].number < sale_willings[0].number:
                a = model.order(buy_willings[0].buyer, sale_willings[0].saler, sale_willings[0].price,
                                buy_willings[0].number)
                orderlist.append(a)
                sale_willings[0].reduce(buy_willings[0].number)
                buy_willings.remove(buy_willings[0])
            else:
                a = model.order(buy_willings[0].buyer, sale_willings[0].saler, sale_willings[0].price,
                                sale_willings[0].number)
                orderlist.append(a)
                buy_willings.remove(buy_willings[0])
                sale_willings.remove(sale_willings[0])
        # 执行订单
        for s in orderlist.copy():
            s.do()
        for s in buy_willings.copy():
            a = model.order(buy_willings[0].buyer, station, station.sale_price, buy_willings[0].number)
            orderlist.append(a)
            a.do()
            buy_willings.remove(s)
        for s in sale_willings.copy():
            a = model.order(station, sale_willings[0].saler, station.sale_price, sale_willings[0].number)
            orderlist.append(a)
            a.do()
            sale_willings.remove(s)
        
        buy_willings = []
        sale_willings = []
        result = []
        for s in citys.copy():
            s.nextday()
            result.append(s.detail())
        station.detail()
        i = 1
        for s in orderlist.copy():
            print("order%s:" % i)
            i += 1  # 用于测试，后期删除
            s.detail()
        
        # 预写入excel
        day = "%s - %s" % (btime.month, btime.day)
        prewrite(result_time, 0, day)  # 写入日期
        prewrite(result_time + 1, 0, "name")  # 写入表头
        prewrite(result_time + 1, 1, "balance")
        prewrite(result_time + 1, 2, "store power")
        prewrite(result_time + 1, 3, "sale power")
        prewrite(result_time + 1, 4, "buy power")
        for s in range(1, len(citys_data) + 1):
            prewrite(result_time + s + 1, 0, result[s - 1][0])
            prewrite(result_time + s + 1, 1, result[s - 1][1])
            prewrite(result_time + s + 1, 2, result[s - 1][2])
            prewrite(result_time + s + 1, 3, result[s - 1][3])
            prewrite(result_time + s + 1, 4, result[s - 1][4])
        prewrite(result_time + len(citys_data) + 2, 0, "order")
        for s in range(1, len(orderlist) + 1):
            prewrite(result_time + len(citys_data) + 2 + s, 0, orderlist[s - 1].buyer.name)
            prewrite(result_time + len(citys_data) + 2 + s, 1, orderlist[s - 1].saler.name)
            prewrite(result_time + len(citys_data) + 2 + s, 2, orderlist[s - 1].number)
        result_time = result_time + len(citys_data) + len(orderlist) + 4
        
        # 善后，模型初始化
        station.buy_power = 0
        station.sale_power = 0
        btime.nextday()
        for s in citys.copy():
            s.buy_power = 0
            s.sale_power = 0
        
        if btime.month == etime.month and btime.day == etime.day:
            break
    
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('My Sheet')
    for s in write_data.copy():
        worksheet.write(s[0], s[1], s[2])
    workbook.save('third result.xls')
