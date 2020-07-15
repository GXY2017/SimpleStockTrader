import io
import time

import pandas as pd
import pywinauto
from pywinauto import clipboard
from pywinauto import keyboard

from .captcha_recognize import captcha_recognize

class SimpleTHSTrader:
    """ 定义一些常量 """

    def __init__(self, exe_path, main_wnd_title):
        """
        exe_path：经纪商软件exe路径
        main_wnd_title：主交易界面标题，inspect 中查找 Name要素。
        """
        self.exe_path = exe_path
        self.title = main_wnd_title
        self.close_client()  # 关闭已有的交易端，避免同一账号重复登录。不关闭多开的登录界面
        self.app = pywinauto.Application().start(self.exe_path, timeout=10)  # 生成后台实例
        self.app.top_window().wait("ready", timeout=10) # 等待登录界面打开
        print("成功启动交易软件。")

    def login(self, id: str, pwd: str):
        # 填入前先删除之前的
        self.app.Dialog.Edit3.set_focus().type_keys('^a{BACKSPACE}') # 通讯码
        self.app.Dialog.ComboBox.Edit.set_focus().type_keys('^a{BACKSPACE}') #账号
        self.app.Dialog.Edit2.set_focus().set_edit_text(u'') # 交易密码

        # 再填写验证码
        if ("中信证券") in self.title:  # 登录时需要填写验证码的
            self.app.Dialog.Edit3.set_focus().type_keys(self.__get_char_login(threshold=175))  # 验证码
        else: #登录无需验证码
            self.app.Dialog.Edit3.set_focus().type_keys(pwd)  # 通讯码
        # 再填写用户名与密码
        self.app.Dialog.ComboBox.Edit.set_focus().type_keys(id)  # 账号
        self.app.Dialog.Edit2.set_focus().type_keys(pwd)  # 交易密码
        self.app.Dialog.child_window(class_name="Button", found_index=0).click()  # 限定子窗口再click()

        try:
            self.app.window(title=self.title).wait("ready",timeout=20)
            print("登录成功！以下关闭各种信息窗口：")
            self.main_wnd = self.app.window(title=self.title)
            self.close_tsxx()  # 关闭所有通知信息
        except:
            print("登录有问题！请查找原因并重新登录。")

    # 关闭提示信息
    def close_tsxx(self):
        '''
        关闭同时弹出的提示窗口。
        1. 提示信息窗口编号都是'#32770'，一下子会跳出来好几个。
        2. app.windows()[0] 等价于 app.window()。
        3. 提示信息窗口都是独立(非内嵌的)的。
        4. 提示信息窗口直接用close()关闭，Y/N选择窗口不行。
        '''
        handles = self.app.windows(class_name='#32770')
        for i in list(range(0, len(handles))):
            print(u'...关闭该窗口: "%s" ' % (handles[i].texts()))
            try:
                handles[i].close()
            except:
                handles[i].type_keys("{ENTER}") # 正在清算
        print(u'\n==> 所有提示信息窗口 已经关闭!')

    # 关闭其他同名交易客户端，self.title
    def close_client(self):
        '''
        关闭交易客户端：
        1. 关闭本地电脑上,已登录的同名交易客户端。只能关闭一个。
        '''
        # self.app.window(title = u"网上股票交易系统5.0").close() # 这个只关闭当前客户端的窗口
        handler = pywinauto.findwindows.find_windows(title=self.title)  # 找到本地电脑上所有同名客户端
        for i in list(range(0, len(handler))):
            print(u'...关闭一个同名客户端:')
            pywinauto.Application().connect(handle=handler[i]).window(title=self.title).close()
            # pywinauto.Application().connect(handle=handler).kill() # 关闭时自动结束进程。

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
        self.main_wnd.child_window(class_name="Button", title="撤买(X)").click()
        self.app.top_window().wait("ready",timeout=1)

        while True:
            """
            分情况处理提示窗口信息
            """
            if ("无可撤委托" or "成功提交" or "提交失败") in self.app.top_window().Static.texts()[0]:
                print(self.app.top_window().Static.texts())
                pywinauto.keyboard.send_keys("{ENTER}")  # 快捷键ENTER = 确定。
                break  # 跳出到return
            elif len(self.app.top_window().Static.texts()[0]) == 0:
                print('没有其他提示，本次操作结束。')
                break
            elif ("回车即可撤单") in self.app.top_window().Static.texts()[0]:
                print('操作提示，无需操作。')
                break
            else:  # 其他情况打印信息后，默认ENTER确认继续
                print(self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                continue

    def cancel_sell_order(self):
        """
        撤买
        """
        self.__select_menu(['撤单[F3]'])
        self.main_wnd.child_window(class_name="Button", title="撤卖(C)").click()
        self.app.top_window().wait("ready",timeout=1)

        while True:
            """
            分情况处理提示窗口信息
            """
            if ("无可撤委托" or "成功提交" or "提交失败") in self.app.top_window().Static.texts()[0]:
                print(self.app.top_window().Static.texts())
                pywinauto.keyboard.send_keys("{ENTER}")  # 快捷键ENTER = 确定。
                break  # 跳出到return
            elif len(self.app.top_window().Static.texts()[0]) == 0:
                print('没有其他提示，本次操作结束。')
                break
            elif ("回车即可撤单") in self.app.top_window().Static.texts()[0]:
                print('操作提示，无需操作。')
                break
            else:  # 其他情况打印信息后，默认ENTER确认继续
                print(self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                continue

    def cancel_all_orders(self):
        """
        撤买
        """
        self.__select_menu(['撤单[F3]'])
        self.main_wnd.child_window(class_name="Button", title="全撤(Z /)").click()
        self.app.top_window().wait("ready",timeout=1)

        while True:
            """
            分情况处理提示窗口信息
            """
            if ("无可撤委托" or "成功提交" or "提交失败") in self.app.top_window().Static.texts()[0]:
                print(self.app.top_window().Static.texts())
                keyboard.send_keys("{ENTER}")  # 快捷键ENTER = 确定。
                break  # 跳出到return
            elif len(self.app.top_window().Static.texts()[0]) == 0:
                print('没有其他提示，本次操作结束。')
                break
            elif ("回车即可撤单") in self.app.top_window().Static.texts()[0]:
                print('操作提示，无需操作。')
                break
            else:  # 其他情况打印信息后，默认ENTER确认继续
                print(self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                continue

    def cancel_by_stock_no(self, stock_no):
        """
        1. 在撤单页面中，输入证券代码后撤单。
        2. 新输入证券代码后，会自动覆盖原有的。每家都要测试。
        """
        self.__select_menu(['撤单[F3]'])
        self.main_wnd.child_window(class_name="Edit").set_edit_text(str(stock_no))  # 自动覆盖原有的
        self.main_wnd.child_window(class_name="Button", title="查询代码").click()
        self.main_wnd.child_window(class_name="Button", title="全撤(Z /)").click()

        # 关闭提示窗口
        while True:
            """
            按窗口信息提示，进行处理
            """
            # 碰到以下情况, 确认退出循环
            if ("无可撤委托") in self.app.top_window().Static.texts()[0]:
                print("有问题：", self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                print('已取消,请检查。')  # 跳出到return
                break
            elif len(self.app.top_window().Static.texts()[0]) == 0:
                print("没有其他提示，本次操作结束。")
                break
            else:  # 其他情况打印信息后，默认ENTER确认继续
                print(self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                continue

    def get_balance(self):
        """
        获取资金情况，这里只是简单地实现。
        """
        self.main_wnd.set_focus()
        self.__select_menu(['查询[F4]', '资金股票'])

        # 关闭各种提示信息
        try:
            self.app.window(class_name="#32770",found_index=0).wait('ready',timeout=5)
            self.close_tsxx()
        except:
            pass

        self.main_wnd.child_window(class_name="Static", found_index=0).wait("ready", timeout= 2)
        # 有啥打印啥。
        i = 0
        while True:
            print(self.main_wnd.child_window(class_name="Static", found_index=i).window_text())
            i += 1
            if "最近查询时间" in self.main_wnd.child_window(class_name="Static", found_index=i).window_text():
                print('查询完毕', self.main_wnd.child_window(class_name="Static", found_index=i).window_text())
                break

    def get_position(self):
        """ 获取持仓 """
        self.main_wnd.set_focus()
        self.__select_menu(['查询[F4]', '资金股票'])
        # 关闭各种提示信息
        try:
            self.app.window(class_name="#32770",found_index=0).wait('ready',timeout=5)
            self.close_tsxx()
        except:
            pass
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

    def get_bills(self,require_dates = True):
        """
        简单实现默认时间段内的对账单功能
        1. 不支持自定义查询日期段
        2. 使用客户端的默认汇总方式
        3. 也没碰到过要翻页的情况
        """
        self.main_wnd.set_focus()
        if "中信证券" in self.title:
            self.__select_menu(['查询[F4]', '对账单'])
        else:
            self.__select_menu(['查询[F4]', '对 账 单'])
        return self.__get_grid_data(require_dates)

    def __trade(self, stock_no, price, amount):
        """
        control_id 要用SPY++ 查到。但很容易变化，尽量不要用。
        """
        self.main_wnd.child_window(control_id=0x408, class_name="Edit").wait('ready', timeout=1)
        self.main_wnd.child_window(control_id=0x408, class_name="Edit").set_text(str(stock_no))  # 设置股票代码
        self.main_wnd.child_window(control_id=0x409, class_name="Edit").type_keys("^a{BACKSPACE}")  # 确保删除原输入
        self.main_wnd.child_window(control_id=0x409, class_name="Edit").set_text(str(price))  # 设置价格
        self.main_wnd.child_window(control_id=0x40A, class_name="Edit").set_text(str(amount))  # 设置股数目
        # 点击卖出or买入
        self.main_wnd.child_window(control_id=0x3EE, class_name="Button").wait('ready',timeout=1)
        self.main_wnd.child_window(control_id=0x3EE, class_name="Button").click()

        while True:
            """
            按窗口信息提示，进行处理
            """
            # 碰到以下情况, 确认退出循环
            if ("失败" or "请输入代码" or "提交失败") in self.app.top_window().Static.texts()[0]:
                print("下单有问题：", self.app.top_window().Static.texts()[0])
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
            elif len(self.app.top_window().Static.texts()[0]) == 0:
                print("没有其他提示，本次操作结束。")
                break
            else:  # 其他情况打印信息后，默认ENTER确认继续
                print(self.app.top_window().Static.texts()[0])
                self.app.top_window().type_keys("{ENTER}")  # 快捷键ENTER = 确定。
                continue

    def __get_grid_data(self, require_dates = False):
        """ GRID数据有时有两个子窗口。一个填日期，一个显示数据。"""
        if require_dates:
            self.main_wnd.child_window(title="Custom1", class_name='CVirtualGridCtrl', found_index=1).set_focus()  # .right_click()  # 模拟右键
        else:
            self.main_wnd.child_window(title="Custom1", class_name='CVirtualGridCtrl', found_index=0).set_focus()
        # 复制
        keyboard.send_keys('^c')  # Ctrl+C
        # 跳出验证码窗口
        self.app.top_window().wait('ready',timeout=2)
        # 验证码错误就重新输入
        trial = 1
        while trial <=5: # 最多尝试5次识别，避免循环
            # 识别并输入验证码，确认关闭验证码窗口
            print("最多识别5次，现在第%d次"%trial)
            self.app.top_window().child_window(class_name='Edit').set_text(self.__get_char())
            self.app.top_window().child_window(class_name='Edit').type_keys("{ENTER}")
            trial += 1
            try:
                "验证码错误" in self.app.top_window().window(class_name="Static",title="验证码错误！！").texts()
            except:
                break #验证码有错，则继续

        data = clipboard.GetData()
        df = pd.read_csv(io.StringIO(data), delimiter='\t', na_filter=False)
        return df.to_dict('records')

    def __get_char(self, threshold=200):
        """
        识别验证码，使用时必须定位到验证码窗口。返回 str
        会将整个窗口截取为图片，所以，窗口不能太大。但child_window(title='提示')
        """
        file_path = "tmp.png"

        self.app.top_window().child_window(control_id=0x965, class_name='Static').\
            capture_as_image().save(file_path)  # 保存验证码
        captcha_num = captcha_recognize(file_path, threshold)  # 识别验证码
        print("captcha result-->", captcha_num)
        return captcha_num

    def __get_char_login(self, threshold=200):
        """
        在登录界面就需要识别验证码。
        threshold: 将200/256转变为白点，对剩下的黑点进行识别。
        """
        file_path = "tmp.png"

        self.app.top_window().set_focus().capture_as_image().save(file_path)  # 保存整个登录界面
        captcha_str = captcha_recognize(file_path, threshold)  # 识别验证码
        captcha_num = [str(s) for s in captcha_str.split() if s.isdigit()][0]  # str(s) 保留00
        print("captcha result-->", captcha_num)
        return captcha_num

    def __select_menu(self, path):
        """ 点击左边菜单 """
        if r"网上股票" or r"网上交易" not in self.main_wnd.window_text():  # 确保在主交易窗口中选择
            self.main_wnd.set_focus()  # 激活平台
            pywinauto.keyboard.send_keys("{ENTER}")
        self.__get_left_menus_handle().get_item(path).click()

    def __get_left_menus_handle(self):
        while True: # 有时需要多次尝试
            try:
                # 在单一账户登录的情况下，同花顺5.0只有一个child_window 是这个名字，不需要control_id
                # login()中已避免多账户登录，或同一账户多次登录。
                handle = self.main_wnd.child_window(class_name='SysTreeView32')
                handle.wait('ready', timeout=2)
                return handle
            except Exception as ex:
                print(ex)
                pass
