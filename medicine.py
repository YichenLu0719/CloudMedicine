
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import requests
import pyodbc
import datetime
import logging
import numpy as np
from bs4 import BeautifulSoup
from smartcard.System import readers


try:
    SelectAPDU = [ 0x00, 0xA4, 0x04, 0x00, 0x10, 0xD1, 0x58, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x11, 0x00 ]

    ReadProfileAPDU = [ 0x00, 0xca, 0x11, 0x00, 0x02, 0x00, 0x00 ]

    # get all the available readers
    r = readers()
    print ("Available readers:", r)

    reader = r[1]
    print ("Using:", reader)

    connection = reader.createConnection()
    connection.connect()

    data, sw1, sw2 = connection.transmit(SelectAPDU)
    print ("Select Applet: %02X %02X" % (sw1, sw2))

    data, sw1, sw2 = connection.transmit(ReadProfileAPDU)

    ID= ''.join(chr(i) for i in data[32:42])
    ID=str(ID)
    print(ID)
except Exception as e:
    print(e)
    print("沒有讀到卡，請關閉再開起健保小綠人並驗pin")
    input("點按enter結束程式")




try:
    url = 'https://medcloud.nhi.gov.tw/imme0008/IMME0008S01.aspx'
    #https://medvpn.nhi.gov.tw/iwse0000/IWSE0001S01.aspx
    service = Service(executable_path = "\chromedriver.exe")
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # 設定為隱藏模式
    driver = webdriver.Chrome(service=service,options=chrome_options)
    driver.get(url)

    try:
        
        locator = (By.ID,'btnClose')  # 定位器
        search_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(locator),
            "找不到指定的元素"
        )
        driver.find_element('id','btnClose').click()
    except:
        pass

    driver.switch_to.frame(0)
    time.sleep(2)
    # 抓取下拉選單元素
    try:
        dropdown = Select(driver.find_element('id', 'ctl00$ContentPlaceHolder1$pg_gvList_input'))
        options = dropdown.options
        num_pages = len(options)
        print(f"下拉選單中有 {num_pages} 頁")
        df = pd.DataFrame()
        # 依序選擇每個選項
        for option_index in range(num_pages):
            # 選擇下拉選單中的選項
            dropdown = Select(driver.find_element('id', 'ctl00$ContentPlaceHolder1$pg_gvList_input'))
            dropdown.select_by_index(option_index)

            # 等待資料載入，可以根據需要自行調整等待時間
            time.sleep(1)

            # 抓取目前頁面的資料        
            source = driver.page_source
            table = pd.read_html(source)
            df_next = pd.DataFrame(table[6])
            driver.switch_to.default_content()
            driver.switch_to.frame(0)
            # 將目前頁面的資料新增到 df 中
            df = pd.concat([df, df_next])

    except NoSuchElementException:
        print("無下拉式選單")
        #driver.switch_to.frame(0)
        source = driver.page_source
        table = pd.read_html(source)
        df = pd.DataFrame(table[6])
    df = df.rename(columns={
            '項次': '項次',
            '來源': '來源',
            '主診斷': '主診斷',
            'ATC5代碼': 'ATC5代碼',
            'ATC3名稱': 'ATC3名稱',
            'ATC5名稱': 'ATC5名稱',
            '複方註記': '複方註記',
            '成分名稱': '成分名稱',
            '藥品 健保代碼': '藥品健保代碼',
            '藥品名稱': '藥品名稱',
            '給藥 日數': '給藥日數',
            '藥品 用量': '藥品用量',
            '藥品規格量': '藥品規格量',
            '用法用量': '用法用量',
            '就醫(調劑)日期(住院用藥起日)': '就醫(調劑)日期(住院用藥起日)',
            '慢連箋 領藥日 (住院用 藥迄日)': '慢連箋領藥日(住院用藥迄日)',
            '單筆 餘藥 日數 試算': '單筆餘藥日數試算',
            '就醫序號': '就醫序號',
            '慢連箋原處方醫事機構代碼': '慢連箋原處方醫事機構代碼',
            '藥品療效 不等': '藥品療效不等',
            '費用年月': '費用年月',
            '新增 過敏資料': '新增過敏資料'
        })
    #df['藥品規格量'] = df['藥品規格量'].fillna(0)
    df = df.fillna('')
    df['就醫(調劑)日期(住院用藥起日)'] = df['就醫(調劑)日期(住院用藥起日)'].str.replace("/", "")
except Exception as e:
    # 設定log檔案名稱
    print(e)
    log_file = "錯誤紀錄.txt"
    # 設定log輸出格式
    log_format = "%(asctime)s - %(levelname)s - %(message)s"
    logging.basicConfig(filename=log_file, level=logging.INFO, format=log_format)
    logging.info(e)
    print("發生錯誤記錄")
    input("點按enter結束程式")



# 建立資料庫連接
server = '' 
database = ''
username = '' 
password = '' 
conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                      'Server=' + server + ';'
                      'Database=' + database + ';'
                      'UID=' + username + ';'
                      'PWD=' + password + ';'
                      'Trusted_Connection=no;')
cursor = conn.cursor()


# In[6]:


# 檢查是否有重複的 ID
count = 0
number = 1
today = datetime.datetime.today().date()
today = today.strftime("%Y%m%d")
try:
    cursor.execute("SELECT COUNT(*) FROM HisDrgTbl WHERE  = ?", (str(ID),))
    counts = cursor.fetchone()[0]
    cursor.execute("SELECT DISTINCT  FROM HisDrgTbl WHERE  = ?", (str(ID),))
    date_result = cursor.fetchall()
    date_result = date_result[0][0].strftime("%Y%m%d")[0:]
    today = datetime.datetime.today().date()
    today = today.strftime("%Y%m%d")
    today_2 = datetime.datetime.today().date()
    today_2= today_2.strftime("%Y-%m-%d")
    #today_SQL=print(date_result[0][0])
        #today = print(today)
        # 如果有重複的 ID，先保留舊資料
    if counts > 0:
        print("重複的資料存在，暫時保留舊資料")
    try:
        # 新增新的資料
        if date_result != today:
            for index, row in df.iterrows():

                    cursor.execute("INSERT INTO HisDrgTbl (, , , , , , , , , , , , , , , , , ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                   (str(ID), number, str(row['來源']), str(row['主診斷']), str(row['ATC5名稱']), str(row['成分名稱']), str(row['藥品健保代碼']), str(row['藥品名稱']), str(row['藥品規格量']), str(row['用法用量']), str(row['就醫(調劑)日期(住院用藥起日)']), str(row['慢連箋領藥日(住院用藥迄日)']), str(row['藥品用量']), str(row['給藥日數']), str(row['單筆餘藥日數試算']), today, "急診下載", str(row['ATC5代碼'])))
                    count += 1
                    number += 1
                    conn.commit()

            print(ID, "資料已新增")
            print(ID, "共", count, "筆資料已全部新增完成")
            if date_result != today:
                cursor.execute("DELETE FROM table WHERE  = ? AND  <> ?", (str(ID), today_2))
                conn.commit()
                print("刪除日期不是今天之重複資料")
            else:
                print("無重複資料")
        else:
            print("此ID今日已新增資料")
    except Exception as e:
        print("新增資料時發生錯誤，取消新增資料:", e)   
    #conn.close()
    #print("程式已執行完畢",ID)
    #input("點按enter結束程式")
except:
        try:
            for index, row in df.iterrows():
                cursor.execute("INSERT INTO table (, , , , , , , , , , , , , , , , , ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                               (str(ID), number, str(row['來源']), str(row['主診斷']), str(row['ATC5名稱']), str(row['成分名稱']), str(row['藥品健保代碼']), str(row['藥品名稱']), str(row['藥品規格量']), str(row['用法用量']), str(row['就醫(調劑)日期(住院用藥起日)']), str(row['慢連箋領藥日(住院用藥迄日)']), str(row['藥品用量']), str(row['給藥日數']), str(row['單筆餘藥日數試算']), today, "急診下載", str(row['ATC5代碼'])))
                count += 1
                number += 1
                conn.commit()
            print(ID, "資料已新增")
            print(ID, "共", count, "筆資料已全部新增完成")
        except Exception as e:
            print("新增資料時發生錯誤，取消新增資料:", e)   
conn.close()
print("程式已執行完畢",ID)
input("點按enter結束程式")





