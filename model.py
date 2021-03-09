import time
import xlrd, xlwt

# 运算
store_cost_pre_day_pre_unit = 10


# 未来需要增加传输损耗和传输花费


class peer:
    def __init__(self, name, get_power, store_power, sale_power, buy_power, speed, sale_price, store_cost, income,
                 balance):
        self.name = name
        self.get_power = get_power  # 每天自然获得的能源
        self.store_power = store_power  # 储存的能源
        self.sale_power = sale_power  # 当天卖出去的能源
        self.buy_power = buy_power  # 当天买进来的能源
        self.speed = speed  # 当天的消耗能源
        self.sale_price = sale_price  # 别的城市从这买能源的单价
        self.store_cost = store_cost  # 每天储存能源的花费
        self.income = income  # 每天卖出的收益
        self.balance = balance  # 还剩下多少钱

    def set_get_power(self, get_power):
        self.get_power = get_power

    def set_speed(self, speed):
        self.speed = speed

    def set_price(self, price):
        self.sale_price = price

    def nextday(self):  # 经过一天
        self.store_power = self.store_power - self.speed + self.get_power  # 更新储存的能源数量
        self.store_cost = self.store_power * store_cost_pre_day_pre_unit  # 更新每天储存的费用
        self.income = self.sale_power * self.sale_price  # 更新每天的收入
        self.balance = self.balance - self.store_cost + self.income  # 更新每天的余额
        # 更新获得的能源
        # 更新速度

    def will_need(self):  # 未来需要购买，1.买刚好够用的 2.一次买多些存起来
        # 暂时！！！！！！写第一种情况
        needbuy_power = self.store_power - self.speed + self.get_power
        if needbuy_power < 0:
            return 0 - needbuy_power
        else:
            return -1

    def will_excess(self):  # 未来会过剩 1.现在卖,保证储量为0 2.留着以后处理（卖还是自己用）
        # 暂时！！！！！写第一种情况
        excess_power = self.store_power - self.speed + self.get_power
        if excess_power > 0:
            return excess_power
        else:
            return -1

    def detail(self):  # 展示结果，把结果写入excel表
        print(self.name)
        print("balance:%srmb" % self.balance)
        print("store power:%skwh" % self.store_power)
        print("sale power:%skwh" % self.sale_power)
        print("buy power:%skwh" % self.buy_power)
        print("")
        return [self.name, self.balance, self.store_power, self.sale_power, self.buy_power]

    def buy_from_station(self, number):
        self.store_power = self.store_power + number
        self.buy_power = self.buy_power + number
        self.balance = self.balance - number * station.sale_price
        station.sale_power = station.sale_power + number

    def sale_to_station(self, number):
        self.store_power = self.store_power - number
        self.sale_power = self.sale_power + number
        self.balance = self.balance + number * station.sale_price
        station.buy_power = station.buy_power + number


class buy_willing:  # 购买欲望 状态1表示为买
    def __init__(self, buyer, number):
        self.buyer = buyer
        self.number = number
        self.statue = 1
        
    def reduce(self, first):  # 购买欲望减少（实际为买了一部分）
        self.number = self.number - first

class sale_willing:  # 卖出欲望 状态0表示为卖
    def __init__(self, saler, number, price):
        self.saler = saler
        self.number = number
        self.price = price
        self.statue = 0

    def reduce(self, first):  # 购买欲望减少（实际为买了一部分）
        self.number = self.number - first

class noWilling:
    def __init__(self,buyer):
        self.buyer = buyer
        self.statue = "null"
        
class order:
    def __init__(self, buyer, saler, price, number):  # 单价和数量
        self.buyer = buyer
        self.saler = saler
        self.price = price
        self.number = number

    def do(self):
        self.buyer.store_power = self.buyer.store_power + self.number
        self.buyer.buy_power = self.buyer.buy_power + self.number
        self.buyer.balance = self.buyer.balance - self.number * self.price

        self.saler.store_power = self.saler.store_power - self.number
        self.saler.sale_power = self.saler.sale_power + self.number
        self.saler.income = self.saler.income + self.number * self.price
        self.saler.balance = self.saler.balance + self.number * self.price

    def detail(self):  # 用于测试，后期删除
        print("buyer:" + self.buyer.name)
        print("saler:" + self.saler.name)
        print("price:%s" % self.price)
        print("number:%s" % self.number)
        print("")


def bubble_sort(nums):  # 冒泡排序
    for i in range(len(nums) - 1):  # 这个循环负责设置冒泡排序进行的次数
        for j in range(len(nums) - i - 1):  # j为列表下标
            if nums[j].price > nums[j + 1].price:
                nums[j], nums[j + 1] = nums[j + 1], nums[j]


station = peer("station", 10000000, 1000000, 0, 0, 0, 50, 0, 0, 10000000000)
if __name__ == '__main__':
    shanghai = peer("shanghai", 100, 0, 0, 0, 500, 10, 0, 0, 100000)
    shanxi = peer("shanxi", 500, 0, 0, 0, 400, 2, 0, 0, 100000)
    xizang = peer("xizang", 1000, 0, 0, 0, 300, 5, 0, 0, 100000)
    station = peer("station", 10000000, 1000000, 0, 0, 0, 50, 0, 0, 10000000000)
    city = [shanghai, shanxi, xizang]

    active = True
    i = 1
    while active:
        localtime = time.asctime(time.localtime(time.time()))
        time.sleep(2)
        print("data:%s" % localtime)
        # 记录所有的购买欲望和卖出欲望
        buy_willings = []  # 保存
        sale_willings = []
        orderlist = []
        for s in city:
            if s.will_need() > 0:
                a = buy_willing(s, s.will_need())  # 发出购买欲望
                buy_willings.append(a)
            elif s.will_excess() > 0:
                a = sale_willing(s, s.will_excess(), s.sale_price)
                sale_willings.append(a)
            else:
                pass
        # 进行匹配
        bubble_sort(sale_willings)  # 把卖出的价格降序
        # 匹配的情况有三种，买大于卖，买等于卖，买小于卖
        active1 = True
        while active1:
            if buy_willings[0].number > sale_willings[0].number:
                a = order(buy_willings[0].buyer, sale_willings[0].saler, sale_willings[0].price,
                          sale_willings[0].number)
                orderlist.append(a)
                buy_willings[0].reduce(sale_willings[0].number)
                sale_willings.remove(sale_willings[0])
            elif buy_willings[0].number < sale_willings[0].number:
                a = order(buy_willings[0].buyer, sale_willings[0].saler, sale_willings[0].price, buy_willings[0].number)
                orderlist.append(a)
                sale_willings[0].reduce(buy_willings[0].number)
                buy_willings.remove(buy_willings[0])
            else:
                a = order(buy_willings[0].buyer, sale_willings[0].saler, sale_willings[0].price,
                          sale_willings[0].number)
                orderlist.append(a)
                buy_willings.remove(buy_willings[0])
                sale_willings.remove(sale_willings[0])
            if buy_willings == [] or sale_willings == []:
                active1 = False
        # 执行订单
        for s in orderlist:
            s.do()
        if buy_willings != []:
            for s in buy_willings:
                s.buyer.buy_from_station(s.number)
        if sale_willings != []:
            for s in sale_willings:
                s.saler.sale_to_station(s.number)
        shanghai.nextday()
        shanxi.nextday()
        xizang.nextday()
        shanghai.detail()
        shanxi.detail()
        xizang.detail()
        station.detail()
        for s in orderlist:  # 用于测试，后期删除
            s.detail()
        for s in city:
            s.buy_power = 0
            s.sale_power = 0
        station.buy_power = 0
        station.sale_power = 0
        i += 1
        if i == 5:
            active = False
            print("finish")

# 购买和出售函数需要修改，大体样式是后面的那种，面对对象编程，购买不应该修改卖家的数据，
# 要注意一天之内多个单子的情况！！！！！！！数据可能出错

# city will need 大于0，发出购买请求，卖方发现请求后，创建订单
# 整个系统看成一个榜，购买意愿贴上去，卖家看上了就选择
