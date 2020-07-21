from china_equity.Simple_ths_trader import SimpleTHSTrader
import pywinauto
from china_equity.captcha_recognize import captcha_recognize
import time
from pywinauto import clipboard
from pywinauto import keyboard

tdx = SimpleTHSTrader(r"C:\cjzq\TdxW.exe",u"通达信网上交易V6")    # 连接客户端

#tdx = pywinauto.Application().start(r"C:\cjzq\TdxW.exe", timeout=10)

mainDlg = tdx.app.window(class_name='#32770')

mainDlg.child_window(class_name="AfxWnd42",found_index = 10, visible_only= True).set_focus()
time.sleep(0.5)
mainDlg.child_window(class_name="AfxWnd42",found_index = 10, visible_only= True).click_input()




tdx.login("55022550","370121")

mainDlg.child_window(class_name = 'ComboBox', found_index = 0).child_window(class_name = 'Edit').set_focus().set_edit_text(u'12345')  # 账号，键盘输入

mainDlg.child_window(class_name = 'SafeEdit',found_index = 0).click()
time.sleep(1)
mainDlg.child_window(class_name='SafeEdit', found_index=0).type_keys('123')  # 密码, type_keys() 之前自动set_focus()


mainDlg.child_window(class_name = 'SafeEdit',found_index = 2).set_focus().type_keys('1234', pause = 0.3)  # 验证码







mainDlg.child_window(class_name="AfxWnd42",found_index = 10, visible_only= True).rectangle()
mainDlg.child_window(class_name = 'SafeEdit',found_index = 2).rectangle()