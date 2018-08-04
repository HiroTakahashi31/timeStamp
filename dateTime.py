import requests
import xlrd
import xlwt
import sys
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# function
def getCellNum(inputNum):
    row = 0
    column = 0
    reList = []

    if inputNum % 7 == 0:
        row = 11 + 20 * ((inputNum // 7) - 1) - 1
        column = 45 - 1
    else:
        row = 11 + 20 * (inputNum // 7) - 1
        column = 3 + (((inputNum % 7) - 1) * 7) - 1

    reList.append(row)
    reList.append(column)
    reList.append(column + 1)
    reList.append(column + 2)
    return reList

# main
args = sys.argv
USER = args[1]
PASS = args[2]
EXCELNAME = args[3]

session = requests.session()

login_info = {
    "client_id":"ksc",
    "email":USER,
    "password":PASS,
    "url":"/employee",
    "login_type":"1"
}

url_login = "https://ssl.jobcan.jp/login/pc-employee"
res = session.post(url_login, data=login_info)
res.raise_for_status()

soup = BeautifulSoup(res.text,"html.parser")
if soup.title.string == "スタッフマイページログイン":
    url_login = "https://id.jobcan.jp/users/sign_in"
    res = session.post(url_login, data=login_info)
    res.raise_for_status()

wb = xlrd.open_workbook(EXCELNAME)
sheet = wb.sheet_by_index(0)
year = sheet.cell_value(4,2)
month = sheet.cell_value(5,2)

if month == 12:
    year = year + 1
    month = 1
else:
    month = month + 1

url_mypage = "https://ssl.jobcan.jp/employee/attendance?list_type=normal&search_type=month&year=" + str(year) + "&month=" + str(month) + "&from%5By%5D=2018&from%5Bm%5D=6&from%5Bd%5D=11&to%5By%5D=2018&to%5Bm%5D=7&to%5Bd%5D=10&type=pdf"

res = session.get(url_mypage)
res.raise_for_status()

soup = BeautifulSoup(res.text,"html.parser")
dates = soup.select(".note tr td > a")
hell = soup.select(".note tr > td:nth-of-type(4)")
heaven = soup.select(".note tr > td:nth-of-type(5)")

outputStr = ""
num = 0

checkDate = dates[0].get_text()
weekday = ['月','火','水','木','金','土','日']

for i,check in enumerate(weekday,start=1):
    if checkDate.find(check) > 0:
        num = i

book = xlwt.Workbook()
s = book.add_sheet("NewSheet1")

for i, (date,cry,joy) in enumerate(zip(dates,hell,heaven),start=num):
    day = date.get_text()
    wordS = cry.get_text().find("(") +1
    startWork = cry.get_text()[wordS:wordS+5]
    wordS = joy.get_text().find("(") +1
    endWork = joy.get_text()[wordS:wordS+5]

    if len(endWork.strip()) > 4: 
        cellList = []
        cellList = getCellNum(i)

        s.write(cellList[0],cellList[1],startWork.replace(":",""))
        s.write(cellList[0],cellList[2],"～")
        s.write(cellList[0],cellList[3],endWork.replace(":",""))

book.save('./AAA.xls')
