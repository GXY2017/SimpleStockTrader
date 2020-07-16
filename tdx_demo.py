from china_equity.Simple_ths_trader import SimpleTHSTrader
import pywinauto
from china_equity.captcha_recognize import captcha_recognize
import time

tdx = SimpleTHSTrader(r"C:\cjzq\TdxW.exe",u"通达信网上交易V6")    # 连接客户端
tdx.login("55022550","370121")


tdx.app.top_window().print_control_identifiers()