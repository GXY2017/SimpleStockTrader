# Purpose of SimpleStockTrader
1. To automate the whole process when trading China onshore stocks, convertible bonds and funds.
2. List of compatible broker systems
    (1) Oriant Securities, 东方证券同花顺独立交易系统
# Requirements
1. Python 3.x but avoid Python 3.7.6 since there is [a problem](https://github.com/pywinauto/pywinauto/issues/867). 
2. Open account with your broker
    (1) Inform your broker of the markets you want to trade, Shanghai A share, Shenzhen A share, ETFs, futures, ETF options, etc..
    (2) Account ID, or Fund ID
    (3) Your password
# Installation
1. install pywinauto and pytesseract
2. set up pytesseract in system environment PATH
3. Copy this py file to your environment
# Examples
1. Log in 
```python
from Simple_ths_trader import SimpleTHSTrader
trader = SimpleTHSTrader(r"C:\东方同花顺独立下单\xiadan.exe") # broker system address
trader.login("ACCOUNTID","PASSWORD")
```
The result
```shell script
登录成功！以下关闭各种信息窗口：
口: "['']" 
关闭该窗口: "['']" 
关闭该窗口: "['']" 
关闭该窗口: "['']" 
关闭该窗口: "['']" 
关闭该窗口: "['']" 
```
2. Send Order
````python
import time
t0 = time.process_time()
trader.sell('600398',6.4,100) # (stock code, price, number of shares)
t1 = time.process_time()
print("使用时间： ", t1-t0)
````
The result
````shell script
股东账号：XXXXXXXX
证券代码：600398
卖出价格：6.400
卖出数量：100
预估金额：634.000
您是否确定以上卖出委托？
提交失败：客户的股票不足。
没有其他提示，本次操作结束。
使用时间：  0.78125
````
# Notices in Coding with Pywinauto
1. Use short-cut keys instead of looking for windows or buttons. 
2. Locate window with window specifications instead of general names.

# Up Coming Tasks
1. connect to more brokers