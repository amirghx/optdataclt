import json
import requests
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import shutil
import xlwings as xw
import pandas as pd


def tzm(pct, date):
    tzm_list = []
    for i in range(len(pct)):
        temp = (30 / date[i]) * pct[i]
        tzm_list.append(temp)
    return tzm_list


def percentage(sell_price, place_price, prices):
    pct_list = []
    for i in range(len(prices)):
        pct = ((int(prices[i]) / (int(sell_price[i]) + int(place_price[i]))) - 1) * 100
        pct_list.append(pct)
    return pct_list


def name_collect(x):
    name_name = []
    for item in x:
        item = item.split(',')
        ISIN = item[26]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        name_name.append(ISIN)
    return name_name


def list_maker_price(std_option_symbol, stock_name, prices):
    price = []
    for j in range(len(std_option_symbol)):
        for i in range(len(stock_name)):
            if std_option_symbol[j] == stock_name[i] or (
                    std_option_symbol[j] == "وغدير1" and stock_name[i] == "وغدیر1") or (
                    std_option_symbol[j] == "كچاد1" and stock_name[i] == "کچاد1") or (
                    std_option_symbol[j] == "فملي1" and stock_name[i] == "فملی1") or (
                    std_option_symbol[j] == "كگل1" and stock_name[i] == "کگل1") or (
                    std_option_symbol[j] == "تاپيكو1" and stock_name[i] == "تاپیکو1"
            ):
                price.append(prices[i])

    return price


def pr_mon(col2, date, buy, sell):
    pr_list = []

    for i in range(len(col2)):
        if (buy[i] == '0' and sell[i] == "0"):
            pr_list.append('-')
        else:
            if date[i] == 0:
                pr_list.append('-')
            else:
                pr = (30 / date[i]) * col2[i]
                pr_list.append(pr)
    return pr_list


def price_collect(x):
    price_name = []
    for item in x:
        item = item.split(',')
        ISIN = item[8]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        price_name.append(ISIN)
    return price_name


def std_name(list):
    compare_list = []
    for item in list:
        name = (item.split('-')[0]).split(' ')[1] + '1'
        compare_list.append(name)
    return compare_list


def shifter(List, title):
    output = [[title]]
    for item in List:
        temp = []
        temp.append(item)
        output.append(temp)
    return output


# find path and create data file
path = os.getcwd()
Created_path = path + '\RawData'

while True:
    # soup stuff

    op = Options()
    op.add_argument('--disable-notifications')
    op.add_experimental_option("prefs", {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    driver = webdriver.Chrome(options=op)

    # Setting Chrome to trust downloads
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': Created_path}}
    command_result = driver.execute("send_command", params)
    driver.implicitly_wait(5)

    # Opening the page
    driver.minimize_window()
    driver.get("https://tse.ir/MarketWatch-ang.html?cat=tradeOption")
    driver.minimize_window()
    time.sleep(10)
    # Click on the button and wait for 10 seconds
    driver.find_element_by_css_selector("a[ng-click='exportTableToExcel()']").click()
    time.sleep(10)
    driver.close()

    filename = max([Created_path + "\\" + f for f in os.listdir(Created_path)], key=os.path.getctime)
    shutil.move(filename, os.path.join(Created_path, r"option.xls"))

    location = Created_path + "\option.xls"
    df = pd.read_excel(location, header=3)
    symbol = df['Unnamed: 0'].tolist()
    Name = df['Unnamed: 1'].tolist()
    sell_price = df['قیمت'].tolist()
    final_sell_price = []

    for price in sell_price:
        temp = price.replace(",", "")
        final_sell_price.append(temp)

    final_buy_price = []
    buy_price = df['قیمت.1'].tolist()
    for price in buy_price:
        temp = price.replace(",", "")
        final_buy_price.append(temp)

    place_price = df['Unnamed: 14'].tolist()
    final_place_price = []

    for price in place_price:
        temp = price.replace(",", "")
        final_place_price.append(temp)

    date = []
    remain_date = df['Unnamed: 13'].tolist()

    for item in Name:
        date.append(item.split('-')[2])

    url = "http://mdapi.tadbirrlc.com/api/Symbol/all"

    response = requests.get(url)
    data = response.text

    parsed = json.loads(data)

    x = parsed['List'].split('],')
    col_one = []
    for i in range(len(sell_price)):
        col_one.append(
            int(final_sell_price[i]) / int(list_maker_price(std_name(Name), name_collect(x), price_collect(x))[i]))
    col_two = []
    for i in range(len(sell_price)):
        col_two.append(((int(final_sell_price[i]) + int(final_place_price[i])) / int(
            list_maker_price(std_name(Name), name_collect(x), price_collect(x))[i])) - 1)
    col_three = []
    for i in range(len(final_buy_price)):
        col_three.append(((int(final_buy_price[i]) + int(final_place_price[i])) / int(
            list_maker_price(std_name(Name), name_collect(x), price_collect(x))[i])) - 1)

    wb = xw.Book('View.xlsx')
    worksheet = wb.sheets('Sheet1')
    worksheet.range('A1').value = shifter(symbol, 'نماد')
    worksheet.range('B1').value = shifter(Name, 'نام')
    worksheet.range('C1').value = shifter(sell_price, 'فروش')
    worksheet.range('D1').value = shifter(buy_price, 'خرید')
    worksheet.range('E1').value = shifter(list_maker_price(std_name(Name), name_collect(x), price_collect(x)),
                                          "قیمت سهم")
    worksheet.range('F1').value = shifter(place_price, 'قیمت اعمال')
    worksheet.range('G1').value = shifter(date, 'سر رسید')
    worksheet.range('H1').value = shifter(remain_date, 'روز تا سر رسید')
    worksheet.range('I1').value = shifter(col_one, 'نسبت بهترین فروش به قیمت لحظه‌ای سهم')
    worksheet.range('J1').value = shifter(col_two, 'درصد پریمیوم دوره برای خرید آپشن')
    worksheet.range('K1').value = shifter(col_three, 'درصد پریمیوم دوره برای فروش آپشن')
    worksheet.range('L1').value = shifter(pr_mon(col_two, remain_date, final_buy_price, final_sell_price),
                                          'درصد پریمیوم ماهانه(خرید)')
    worksheet.range('M1').value = shifter(pr_mon(col_three, remain_date, final_buy_price, final_sell_price), 'درصد '
                                                                                                             'پریمیوم'
                                                                                                             'ماهانه('
                                                                                                             'فروش)')
    print('done')
    time.sleep(120)
