import pywinauto
from pywinauto import clipboard
from pywinauto import keyboard
from .captcha_recognize import captcha_recognize
import pandas as pd
import io
import time
"""测试"""

class SimpleTHSTrader:
    """ 定义一些常量 """

    def __init__(self, exe_path):
        """此处增加自动登录功能"""
        self.exe_path = exe_path
        self.title = u'网上股票交易系统5.0' # 用inspect查找主交易界面的名称，不是登录界面
        print("成功创建一个新应用实例。")

    def login(self, id:str, pwd:str):
        self.app = pywinauto.Application().start(self.exe_path, timeout=10) #打开登录界面
        self.close_client() # 关闭已有的交易端（已进入交易界面）
        self.app.Dialog.ComboBox.Edit.set_edit_text(id)  # 资金账户
        self.app.Dialog.Edit2.set_edit_text(pwd)  # 交易密码
        self.app.Dialog.Edit3.set_edit_text(pwd)  # 通信密码
        self.app.Dialog.child_window(class_name="Button", found_index=0).click()  # 限定子窗口再click()
        time.sleep(10) # 等待窗口打开
        if self.app.window(title = self.title):
            print("登录成功！以下关闭各种信息窗口：")
            self.main_wnd = self.app.window(title = self.title)
            time.sleep(3)  # 等待窗口打开, 不然会提示“未收到服务器”之类的
            self.close_tsxx()  # 关闭所有通知信息
        else:
            print("登录有问题，请自行查找原因！")

    # 关闭提示信息
    def close_tsxx(self):
        '''
        1. 提示信息窗口编号都是'#32770'，一下子会跳出来好几个。
        2. app.windows()[0] 等价于 app.window()。
        3. 提示信息窗口都是独立(非内嵌的)的。
        4. 提示信息窗口直接用close()关闭，Y/N选择窗口不行。
        '''
        handles = self.app.windows(class_name='#32770')
        for i in list(range(0,len(handles))):
            print(u'...关闭该窗口: "%s" ' % (handles[i].texts()))
            handles[i].close()
        print(u'\n==> 所有提示信息窗口 已经关闭!')

    # 关闭其他同名交易客户端，self.title
    def close_client(self):
        '''
        关闭交易客户端：
        1. 关闭本地电脑上,已登录的同名交易客户端。只能关闭一个。
        '''
        #self.app.window(title = u"网上股票交易系统5.0").close() # 这个只关闭当前客户端的窗口
        handler = pywinauto.findwindows.find_windows(title= self.title) # 找到本地电脑上所有同名客户端
        for i in list(range(0, len(handler))):
            print(u'...关闭一个同名客户端:')
            pywinauto.Application().connect(handle=handler[i]).window(title= self.title).close()
            #pywinauto.Application().connect(handle=handler).kill() # 关闭时自动结束进程。

    def buy(self, stock_no, price, amount):
        """ 买入 """
        self.main_wnd.set_focus()
        self.__select_menu(['买入[F1]'])
        return self.__trade(stock_no, price, amount)

    def sell(self, stock_no, price, amount):
        """ 卖出 """
        self.main_wnd.set_focus()
        self.__select_menu(['卖出[F2]'])
        return self.__trade(stock_no, price, amount)

    def cancel_buy_order(self):
        """
        撤买
        """
        self.__select_menu(['撤单[F3]'])
        self.main_wnd.child_window(class_name = "Button", title = "撤买(X)").click()
        keyboard.send_keys("{ENTER}") # 快捷键
        time.sleep(0.2)

        while True:
            """
            分情况处理提示窗口信息
            """
            if ("无可撤委托" or "成功提交" or "提交失败") in self.app.top_window().Static.texts()[0]:
                print(self.app.top_window().Static.texts())
                pywinauto.keyboard.send_keys("{ENTER}")  # 快捷键ENTER = 确定。
                time.sleep(1)
                break  # 跳出到return
            elif len(self.app.top_window().Static.texts()[0])==0:
                print('没有其他提示，本次操作结束。')
                break
            else: # 其他情况打印信息后，默认ENTER确认继续
                print(self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                continue

    def cancel_sell_order(self):
        """
        撤买
        """
        self.__select_menu(['撤单[F3]'])
        self.main_wnd.child_window(class_name = "Button", title = "撤卖(C)").click()
        keyboard.send_keys("{ENTER}") # 快捷键
        time.sleep(0.2)

        while True:
            """
            分情况处理提示窗口信息
            """
            if ("无可撤委托" or "成功提交" or "提交失败") in self.app.top_window().Static.texts()[0]:
                print(self.app.top_window().Static.texts())
                pywinauto.keyboard.send_keys("{ENTER}")  # 快捷键ENTER = 确定。
                time.sleep(1)
                break  # 跳出到return
            elif len(self.app.top_window().Static.texts()[0])==0:
                print('没有其他提示，本次操作结束。')
                break
            else: # 其他情况打印信息后，默认ENTER确认继续
                print(self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                continue

    def cancel_all_orders(self):
        """
        撤买
        """
        time.sleep(0.5)
        self.__select_menu(['撤单[F3]'])
        self.main_wnd.child_window(class_name = "Button", title = "全撤(Z /)" ).click()
        keyboard.send_keys("{ENTER}") # 快捷键
        time.sleep(1)

        while True:
            """
            分情况处理提示窗口信息
            """
            if ("无可撤委托" or "成功提交" or "提交失败") in self.app.top_window().Static.texts()[0]:
                print(self.app.top_window().Static.texts())
                pywinauto.keyboard.send_keys("{ENTER}")  # 快捷键ENTER = 确定。
                time.sleep(1)
                break  # 跳出到return
            elif len(self.app.top_window().Static.texts()[0])==0:
                print('没有其他提示，本次操作结束。')
                break
            else: # 其他情况打印信息后，默认ENTER确认继续
                print(self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                continue

    def cancel_by_stock_no(self):
        print('this is under construction.')

    def get_balance(self):
        """
        获取资金情况，这里只是简单地实现。
        """
        self.main_wnd.set_focus()
        self.__select_menu(['查询[F4]', '资金股票'])
        self.close_tsxx() # 关闭各种提示信息
        self.main_wnd.child_windows(class_name="Static").wait('ready')
        # 有啥打印啥。
        i = 0
        while True:
            if not self.main_wnd.child_window(class_name="Static", found_index=i).exists():
                break
            else:
                print(self.main_wnd.child_window(class_name="Static", found_index=i).window_text())
                i += 1

    def check_trade_finished_not_good(self, order_no):
        """
        判断订单是否完成
        用合同编号非常不直观。
        """
        self.__select_menu(['卖出[F2]'])
        time.sleep(1)
        self.__select_menu(['撤单[F3]'])
        cancelable_orders = self.__get_grid_data(is_order_sent = True)
        #        print(cancelable_orders)
        for i, order in enumerate(cancelable_orders):
            if str(order["合同编号"]) == str(order_no):  # 如果订单未完成，就意味着可以撤单
                if order["成交数量"] == 0:
                    return False
        return True

    def get_position(self):
        """ 获取持仓 """
        self.main_wnd.set_focus()
        self.__select_menu(['查询[F4]', '资金股票'])
        self.close_tsxx()  # 关闭提示信息
        return self.__get_grid_data()

    def get_today_order_sent(self):
        """ 获取当日委托单 """
        self.main_wnd.set_focus()
        self.__select_menu(['查询[F4]', '当日委托'])
        return self.__get_grid_data()

    def get_today_trades(self):
        self.main_wnd.set_focus()
        self.__select_menu(['查询[F4]', '当日成交'])
        return self.__get_grid_data()

    def get_bills(self):
        """
        简单实现默认时间段内的对账单功能
        1. 不支持自定义查询日期段
        2. 使用客户端的默认汇总方式
        3. 也没碰到过要翻页的情况
        """
        self.main_wnd.set_focus()
        self.__select_menu(['查询[F4]', '对 账 单'])
        return self.__get_grid_data()

    def __trade(self, stock_no, price, amount):
        """
        control_id 要用SPY++ 查到。但很容易变化，尽量不要用。
        """
        self.main_wnd.child_window(control_id=0x408, class_name="Edit").wait('ready')
        self.main_wnd.child_window(control_id=0x408, class_name="Edit").set_text(str(stock_no))  # 设置股票代码
        self.main_wnd.child_window(control_id=0x409, class_name="Edit").type_keys("^a{BACKSPACE}") # 确保删除原输入
        self.main_wnd.child_window(control_id=0x409, class_name="Edit").set_text(str(price))  # 设置价格
        self.main_wnd.child_window(control_id=0x40A, class_name="Edit").set_text(str(amount))  # 设置股数目
        # 点击卖出or买入
        self.main_wnd.child_window(control_id=0x3EE, class_name="Button").wait('ready')
        self.main_wnd.child_window(control_id=0x3EE, class_name="Button").click()

        while True:
            """
            按窗口信息提示，进行处理
            """
            # 碰到以下情况, 确认退出循环
            if ("失败" or "请输入代码" or "提交失败") in self.app.top_window().Static.texts()[0]:
                print("下单有问题：",self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                print('已取消下单,请检查。')  # 跳出到return
                break
            # 碰到以下情况，选择否，ALT+N 或 Ctrl+n，然后退出循环
            elif "超出涨跌停限制" in self.app.top_window().Static.texts()[0]:
                # 超限时，所有都点完后，系统会自动撤单
                print("检测涨跌停：", self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("%n")  # 快捷键ALT+N = 否。
                print('超出涨跌停限制，已取消下单')  # 跳出到return
                break
            elif len(self.app.top_window().Static.texts()[0])==0:
                print("没有其他提示，本次操作结束。")
                break
            else: # 其他情况打印信息后，默认ENTER确认继续
                print(self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                continue


    def __get_grid_data(self, is_order_sent= False):
        """ 获取grid里面的数据 """
        self.main_wnd.child_window(title="Custom1", class_name='CVirtualGridCtrl').wait('ready')
        self.main_wnd.child_window(title="Custom1", class_name='CVirtualGridCtrl').set_focus()#.right_click()  # 模拟右键
        # 复制Grid
        if not is_order_sent: # 确认不是针对已发出委托
            pywinauto.keyboard.send_keys('^c')  # Ctrl+C
            time.sleep(0.1)
        while True:
            # 识别并输入验证码，确认关闭验证码窗口
            self.app.top_window().child_window(control_id=0x964, class_name='Edit').set_text(self.__get_char())
            self.app.top_window().set_focus()
            keyboard.send_keys("{ENTER}")
            try:
                print(self.app.top_window().child_window(control_id=0x966, class_name='Static').child_window_text())
            except:
                break
            time.sleep(0.1)

        data = clipboard.GetData()
        df = pd.read_csv(io.StringIO(data), delimiter='\t', na_filter=False)
        return df.to_dict('records')

    def __get_char(self):
        """
        识别验证码，使用时必须定位到验证码窗口。返回 str
        会将整个窗口截取为图片，所以，窗口不能太大。但child_window(title='提示')
        """
        file_path = "tmp.png"

        self.app.top_window().child_window(control_id=0x965, class_name='Static'). \
            capture_as_image().save(file_path)  # 保存验证码
        captcha_num = captcha_recognize(file_path)  # 识别验证码
        print("captcha result-->", captcha_num)
        return captcha_num

    def __select_menu(self, path):
        """ 点击左边菜单 """
        if r"网上股票" or r"网上交易" not in self.main_wnd.window_text(): # 确保在主交易窗口中选择
            self.main_wnd.set_focus()  # 激活平台
            pywinauto.keyboard.send_keys("{ENTER}")
        self.__get_left_menus_handle().get_item(path).click()

    def __get_left_menus_handle(self):
        while True:
            try:
                # 在单一账户登录的情况下，同花顺5.0只有一个child_window 是这个名字，不需要control_id
                # login()中已避免多账户登录，或同一账户多次登录。
                handle = self.main_wnd.child_window(class_name='SysTreeView32')
                handle.wait('ready', 2)  # sometime can't find handle ready, must retry
                return handle
            except Exception as ex:
                print(ex)
                pass

