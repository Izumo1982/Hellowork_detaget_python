from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import openpyxl
import urllib
import os
import pandas as pd
import s_mail

from datetime import datetime as dt

def conma_out(str):
	str= str.replace('\xa0', '').replace('\n', '  ').replace('\t', ' ').replace(',', ';')
	return str
    
def hallow_w():
    tdatetime = dt.now()
    tstr = tdatetime.strftime('%Y%m%d') + ","
    tstr_start = tdatetime.strftime('%Y-%m-%d %H:%M:%S')

    f_name = "brandnew.xlsx"
    book = openpyxl.load_workbook(f_name)
    sheet = book.worksheets[0]
    cell = sheet['A2']
    brnew_data = []
    for row in sheet.rows:
            brnew_data.append(row[1].value.replace('\n', ''))
            brnew_date = row[0].value

    url = "https://www.hellowork.mhlw.go.jp/"
    url2 = "https://www.hellowork.mhlw.go.jp/kensaku/"
    print(str(brnew_date))

    df = pd.read_csv("kyujin_ravel.csv")
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(ChromeDriverManager().install())

    driver.get(url)
    time.sleep(1)

    driver.find_element_by_class_name("retrieval_icn").click()
    time.sleep(1)

    element = driver.find_element_by_id("ID_kSNoJo")
    element.send_keys("00000")#実際は自分の求職者番号を入れています
    time.sleep(1)

    element = driver.find_element_by_id("ID_kSNoGe")
    element.send_keys("00000000")#実際は自分の求職者番号を入れています
    time.sleep(1)

    element = driver.find_element_by_id("ID_nenreiInput")
    element.send_keys("38")
    time.sleep(1)

    element = driver.find_element_by_id("ID_sKGYBRUIJo1")
    element.send_keys("102")
    time.sleep(1)


    element = driver.find_element_by_id("ID_sKGYBRUIJo2")
    element.send_keys("104")
    time.sleep(1)

    element = driver.find_element_by_id("ID_tDFK1CmbBox")
    Select(element).select_by_value("32")
    time.sleep(1)

    driver.find_element_by_id("ID_Btn").click()
    time.sleep(1)

    element = driver.find_element_by_id("ID_rank1CodeMulti")
    Select(element).select_by_value("32201")
    Select(element).select_by_value("32203")
    Select(element).select_by_value("32209")
    Select(element).select_by_value("32343")
    time.sleep(1)

    driver.find_element_by_id("ID_ok").click()
    time.sleep(1)

    driver.find_element_by_id("ID_searchBtn").click()
    time.sleep(1)

    element = driver.find_element_by_id("ID_fwListNaviDispBtm")
    Select(element).select_by_value("50")
    time.sleep(1)

    soup = BeautifulSoup(driver.page_source, "html.parser")

    f_n = 0
    url_hai = []
    for element in soup.find_all("a"):
            urll = element.get("href")
            if urll[0:62] == "./GECA110010.do?screenId=GECA110010&action=dispDetailBtn&kJNo=":
                    urll = urllib.parse.urljoin(url2,urll)
                    url_hai.append(urll)
                    f_n += 1
    driver.close()
    f_n = 0
    flg = ""
    for t_url in url_hai:
            rec = [''] * len(df.columns)
            rec[0] = tstr_start
            print(f_n)
            driver = webdriver.Chrome(ChromeDriverManager().install())
            driver.get(t_url)
            time.sleep(1)
            soup = BeautifulSoup(driver.page_source, "html.parser")
            f_n += 1
            for element in soup.find_all("tr"):
                    if element.find("th") is not None:
                            insatu1 = element.find("th").text
                            insatu1 = conma_out(insatu1)
                            try:
                                    j = df.columns.get_loc(insatu1)
                            except:
                                    j = 999
                    if element.find("td") is not None:
                            insatu2 = element.find("td").text
                            if j == 999:
                                    rec[len(df.columns) - 1] = insatu2
                            else:
                                    rec[j] = insatu2

            if len(brnew_data) <= f_n :
                    flg = flg + str(f_n) + "件目　元データなし\n"
                    print(flg + str(f_n) + "件目　元データなし\n")
            elif rec[1].replace('\n', '') != brnew_data[f_n]:
                    flg = flg + str(f_n) + "件目　データ更新\n"
                    print(str(f_n) + "件目　データ更新")

            df.loc[f_n] = rec
            driver.close()

    if flg == "":
            print(str(brnew_date) + "と同一内容")            
            flg = brnew_date + "と同一内容"            
    else:
            print(flg)
    
    df.to_excel(tstr + "102-104.xlsx", index=False)
    df.to_excel("brandnew.xlsx", index=False)
    s_mail.sending_mail("send@xxxx.com", "recieve@xxxx.com", "求人情報更新", flg)
    #実際は自分のアドレスを入力しています
