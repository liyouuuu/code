import time
import xlrd
import xlwt
import model


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


def prewrite(row, col, num):
    a = [row, col, num]
    write_data.append(a)


if __name__ == '__main__':
    btime = data(0, 0, 0)
    etime = data(0, 0, 0)
    # 设置开始时间和结束时间
    btime.setday(1)
    btime.setmonth(1)
    etime.setday(1)
    etime.setmonth(11)
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
    
    # 分开记录城市的willings
    citys_willings = []
    i = 0
    for s in citys.copy():
        city_willings = []
        for b in range(btime.month, etime.month):
            if s.will_need() > 0:
                a = model.buy_willing(s, s.will_need())  # 发出购买欲望
                city_willings.append(a)
            elif s.will_excess() > 0:
                a = model.sale_willing(s, s.will_excess(), s.sale_price)
                city_willings.append(a)
            else:
                pass
            s.set_speed(citys_data[i][b][5])
            s.set_get_power(citys_data[i][b][1])
        i += 1
        citys_willings.append(city_willings)
    print(citys_willings[1][1].number)
