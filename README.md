# Purpose of SimpleStockTrader  
1. To automate the whole process when trading China onshore stocks, convertible bonds, funds and repos.  
2. This project focuses on user login and order execution. It excludes market data and strategies.  
3. List of connected broker trading platform  
    (1) Orient Securities, 东方证券同花顺独立交易系统  
    (2) Citic Securities, 中信证券至胜全能版独立交易系统
# Requirements  
1. Python 3.x but please avoid using Python 3.7.6 since there is [a problem](https://github.com/pywinauto/pywinauto/issues/867).   
2. Better to use python 32 bit. 64 bit Python you will get a warning.  
3. Open account with your broker  
    (1) Inform your broker of the markets you want to trade, Shanghai A share, Shenzhen A share, ETFs, futures, ETF options, etc..    
    (2) Account ID, or Funding Account ID  
    (3) Your password  
# Installation  
1. install pywinauto and pytesseract  
2. set up pytesseract in system environment PATH  
3. Copy this py file to your environment  
# Available functions  
1. Log in  
2. Send order  
3. Query balance and holdings  
4. Query transactions  
5. Cancel orders by direction and stock code  

# Problems Unsolved  
1. Pywinauto mimics user actions, it will pull app to the front screen when running.  
2. Not able to listen to the responses from brokers if the sent order is done.  

# Tips to Speed Up A Little with Pywinauto  
1. Use short-cut keys instead of looking for windows or buttons.  
2. Locate window with window specifications instead of general titles or types.  
3. Use window().wait('ready',timeout) instead of time.sleep() when waiting. 

# What's Next  
1. connect to more brokers  
- Tdx (通达信) seems to be anti-automation. I tried [JinChangJiang 金长江财智版V11.82](https://www.95579.com/main/software/index.shtml) 
,which is based on Tdx system. I cannot type_keys() to password and account windows.  

# Even More
1. Github [China-quant](https://github.com/china-quant) is a good reference. 
2. Here is an archived project two years ago, [TdxTradeServer](https://github.com/rainx/TdxTradeServer), 
which is based on trade.dll file. Obviously, it goes deeper and runs faster. This is a C++ project.  
   This maybe a good direction to avoid operating via GUI.  
3. [Automate trade module in Tdxw](https://www.cnblogs.com/duan-qs/p/10296462.html) gives a good example to automate
tqd platform. It is a python 2.x project.  

# Useful Tools    
1. Inspect installed with Visual Studio    
2. Accessbility Insights installed with windows SDK, "entire app" mode    
3. Spy++ installed with windows SDK    

# Examples
1. Log in 
```python
import SimpleTHSTrader
trader = SimpleTHSTrader(r"C:\东方同花顺独立下单\xiadan.exe") # broker system address
trader.login("ACCOUNTID","PASSWORD")
```
If you need two instances for one account, you can change codes in login() function to achieve this. 
I don't think this is necessary.          
The result:
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
trader.buy('128036',100,10) # (stock code, price, number of shares)
t1 = time.process_time()
print("使用时间： ", t1-t0)
````
The result:
````shell script
股东账号：XXXXXXXX
证券代码：128036
买入价格：100.000
买入数量：10
您是否确定以上买入委托？
提交失败：委托在非交易时间。
没有其他提示，本次操作结束。
使用时间：  0.390625
````



