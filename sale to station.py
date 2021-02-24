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
        for s in [1, 3, 5, 7, 8, 10, 12]:
            if s == self.month:
                self.length = 31
        for s in [4, 6, 9, 11]:
            if s == self.month:
                self.length = 30
        if self.month == 2:
            self.length = 28

    def setday(self, day):
        self.day = day

    def nextmonth(self):
        if self.month < 12:
            self.month += 1
        else:
            self.month = 1
        for s in [1, 3, 5, 7, 8, 10, 12]:
            if s == self.month:
                self.length = 31
        for s in [4, 6, 9, 11]:
            if s == self.month:
                self.length = 30
        if self.month == 2:
            self.length = 28

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
    btime.setday(31)
    btime.setmonth(1)
    etime.setday(2)
    etime.setmonth(2)
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
        city = model.peer(s[btime.month][10], s[btime.month - 1][1], s[btime.month - 1][2], s[btime.month - 1][3],
                          s[btime.month - 1][4], s[btime.month - 1][5], s[btime.month - 1][6],
                          s[btime.month - 1][7], s[btime.month - 1][8], s[btime.month - 1][9])
        citys.append(city)
    while active:
        time.sleep(2)
        print("data: %s-%s" % (btime.month, btime.day))
        i = 0
        for s in citys:
            s.set_get_power(citys_data[i][btime.month - 1][1])
            s.set_speed(citys_data[i][btime.month - 1][5])
            s.set_price(citys_data[i][btime.month - 1][6])
            i += 1
        # 多出来的能源都卖给发电站，少的都从发电站买
        buy_willings = []
        sale_willings = []
        for s in citys.copy():
            if s.will_need() > 0:
                a = model.buy_willing(s, s.will_need())
                buy_willings.append(a)
            elif s.will_excess() > 0:
                a = model.sale_willing(s, s.will_excess(), s.sale_price)
                sale_willings.append(a)
            else:
                pass
        if buy_willings != []:
            for s in buy_willings.copy():
                s.buyer.buy_from_station(s.number)
        if sale_willings != []:
            for s in sale_willings.copy():
                s.saler.sale_to_station(s.number)

        result = []
        for s in citys.copy():
            s.nextday()
            result.append(s.detail())
        station.detail()
        i = 1

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
        result_time = result_time + len(citys_data) + 3

        # 善后，模型初始化
        station.buy_power = 0
        station.sale_power = 0
        btime.nextday()
        for s in citys.copy():
            s.buy_power = 0
            s.sale_power = 0

        # 结束，退出循环
        if btime.month == etime.month and btime.day == etime.day:
            active = False
            print("finish")

    # 写入excel
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('My Sheet')
    for s in write_data:
        worksheet.write(s[0], s[1], s[2])
    workbook.save('to station result.xls')
