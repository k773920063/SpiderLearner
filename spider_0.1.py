import urllib.request
import requests
import datetime
import pymysql
import json
import xlrd
import re
import os


def Make_Data_List():
    now_time = datetime.datetime.now()
    i = 0
    list = []
    for num in range(0, -365, -1):
        the_day = now_time + datetime.timedelta(days=num)
        list.append(the_day.strftime('%Y-%m-%d'))
    return list


def getHtml(url):
    html = requests.get(url)
    return html


# 获取融资余额 融券余额 收盘股价
def Re_Get_Data(Book, date):
    Cache = re.search(date, Book)
    if Cache:
        Loc_Start = Cache.end()
        # 获取融资余额
        pattern = re.compile(r'"rzye":')
        loc_start = pattern.search(Book, Loc_Start).end()
        if(Book[loc_start + 1] == '-'):
            Rzye = None
        else:
            pattern = re.compile(r"\d+\.?\d*")
            Rzye = pattern.search(Book, loc_start).group()

        # 获取融券余额
        pattern = re.compile(r'"rqye":')
        loc_start = pattern.search(Book, Loc_Start).end()
        if (Book[loc_start + 1] == '-'):
            Rqye = None
        else:
            pattern = re.compile(r"\d+\.?\d*")
            Rqye = pattern.search(Book, loc_start).group()

        # 获取收盘股价
        pattern = re.compile(r'"close":')
        loc_start = pattern.search(Book, Loc_Start).end()
        if (Book[loc_start + 1] == '-'):
            price = None
        else:
            pattern = re.compile(r"\d+\.?\d*")
            price = pattern.search(Book, loc_start).group()

        return Rzye, Rqye, price
    else:
        return None

def Mk_new_table(stock_num):
    sql = "create table Stock_" + str(stock_num) + " (Stock_Date date not null,Price double,Rzye double,Rqye double)"
    cursor.execute(sql)


def insert_data_toDB(stock_num, date, price, Rzye, Rqye):
    # db = pymysql.connect('localhost','Pyer','pythonpass','test1')
    # cursor = db.cursor()
    if not price:
        price = 'null'
    if not Rzye:
        Rzye = 'null'
    if not Rqye:
        Rqye = 'null'
    sql = "insert into Stock_" + str(stock_num) + " values(\"" + date + "\"," + str(price) + "," + str(Rzye) + "," + str(Rqye) + ")"
    cursor.execute(sql)
    db.commit()


def read_stock_list():
    workbook = xlrd.open_workbook(r'D:\list.xls')
    #print(workbook.sheet_names())

    sheet1 = workbook.sheet_by_index(0)

    col1 = sheet1.col_values(0)
    col2 = sheet1.col_values(1)

    return col1, col2


stock_num_list = []
cache = read_stock_list()
stock_num_list = cache[0]
Date_List = Make_Data_List()

db = pymysql.connect('localhost','Pyer','pythonpass','test1')
cursor = db.cursor()
count = 0
for i in range(0,len(stock_num_list)):
    Input_code = stock_num_list[i]
    Input_url = 'http://dcfm.eastmoney.com/em_mutisvcexpandinterface/api/js/get?type=RZRQ_DETAIL_NJ&token=70f12f2f4f091e459a279469fe49eca5&filter=(scode=%27' + Input_code + '%27)&st={sortType}&sr={sortRule}&p={page}&ps={pageSize}&js=var%20{jsname}={pages:(tp),data:(x)}{param}'
    Book = getHtml(Input_url).text
    Mk_new_table(stock_num_list[i])
    for j in range(0,365):
        Cache = Re_Get_Data(Book,Date_List[j])
        if Cache:
            insert_data_toDB(Input_code,Date_List[j],Cache[2],Cache[0],Cache[1])
    print("Finished:"+str(stock_num_list[i])+ "  total:"+str(i+1)+"/525")
db.close()
