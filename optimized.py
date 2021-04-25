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


# def test(city_willings):
#     for i in range(btime.month - 1, etime.month - 2):
#         if city_willings[i].statue != 0:
#             return 0
#         else:
#
def PreWrite(row, col, num):
    a = [row, col, num]
    thirdresult.append(a)


btime = data(0, 0, 0)
etime = data(0, 0, 0)
# 设置开始时间和结束时间
btime.setday(1)
btime.setmonth(1)
etime.setday(1)
etime.setmonth(12)
if __name__ == '__main__':
    
    # 设置预见的天数
    foresee = 4
    #
    thirdresult = []
    write_data = []
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
    print(citys[0].name)
    print(citys[0].speed)
    # 分开记录城市的willings
    CitysWillings = []
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
                a = model.noWilling(s)
                city_willings.append(a)
            s.set_speed(citys_data[i][b][5])
            s.set_get_power(citys_data[i][b][1])
        i += 1
        CitysWillings.append(city_willings)
    # print(citys_willings[1][1].number) 第二个城市的第二个月的愿望的数目
    
    # 打印出原始的消费行为
    prewrite(0, 0, "time")  # 写入表头
    prewrite(0, 1, "willing")
    prewrite(0, 2, "peer")
    prewrite(0, 3, "number")
    prewrite(0, 4, "price")
    j = 0
    for s in CitysWillings.copy():
        for i in range(0, etime.month - btime.month):
            if s[i].statue == 0:
                prewrite(i + 1 + j, 0, i + 1)
                prewrite(i + 1 + j, 1, "sale willing")
                prewrite(i + 1 + j, 2, s[i].saler.name)
                prewrite(i + 1 + j, 3, s[i].number)
                prewrite(i + 1 + j, 4, s[i].price)
            elif s[i].statue == 1:
                prewrite(i + 1 + j, 0, i + 1)
                prewrite(i + 1 + j, 1, "buy willing")
                prewrite(i + 1 + j, 2, s[i].buyer.name)
                prewrite(i + 1 + j, 3, s[i].number)
            else:
                prewrite(i + 1 + j, 0, i + 1)
                prewrite(i + 1 + j, 1, "no willing")
                prewrite(i + 1 + j, 2, s[i].buyer.name)
                prewrite(i + 1 + j, 3, 0)
        j = j + len(s)
    
    # 最后一行 print(write_data[-1][0])
    new = write_data[-1][0]
    # 整合消费行为
    for s in CitysWillings:
        for i in range(0, etime.month - btime.month):
            if s[i].statue == 0:
                for j in range(1, 1 + foresee):  # 预判4天
                    try:
                        if s[i + j].statue == 1:
                            s[i].saleAndBuy(s[i + j])
                        if s[i].number == 0:
                            break
                    except BaseException:
                        break
    prewrite(0, 6, "optimized")
    prewrite(0, 7, "time")  # 写入表头
    prewrite(0, 8, "willing")
    prewrite(0, 9, "peer")
    prewrite(0, 10, "number")
    prewrite(0, 11, "price")
    j = 0
    for s in CitysWillings.copy():
        for i in range(0, etime.month - btime.month):
            if s[i].statue == 0:
                prewrite(i + 1 + j, 7, i + 1)
                prewrite(i + 1 + j, 8, "sale willing")
                prewrite(i + 1 + j, 9, s[i].saler.name)
                prewrite(i + 1 + j, 10, s[i].number)
                prewrite(i + 1 + j, 11, s[i].price)
            elif s[i].statue == 1:
                prewrite(i + 1 + j, 7, i + 1)
                prewrite(i + 1 + j, 8, "buy willing")
                prewrite(i + 1 + j, 9, s[i].buyer.name)
                prewrite(i + 1 + j, 10, s[i].number)
            else:
                prewrite(i + 1 + j, 7, i + 1)
                prewrite(i + 1 + j, 8, "no willing")
                prewrite(i + 1 + j, 9, s[i].buyer.name)
                prewrite(i + 1 + j, 10, 0)
        j = j + len(s)
    
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('My Sheet')
    for s in write_data.copy():
        worksheet.write(s[0], s[1], s[2])
    workbook.save('optimized.xls')
