import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from time import sleep
import pyautogui
import time
from openpyxl import load_workbook

url = "https://rozetka.com.ua/notebooks/c80004/"
fn = "Data.xlsx"
wb = load_workbook(fn)
ws = wb["Sheet"]

url = "https://rozetka.com.ua/notebooks/c80004/"

options = webdriver.ChromeOptions()
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36")
options.headless = True
driver = webdriver.Chrome(executable_path=r"C:\Program\Projects\. Learning\Parcing\Selenium learning\chromedriver.exe", options=options)


list_urls = []


for pages in range(9):
    page = pages + 1

    driver.get(url=f"https://rozetka.com.ua/notebooks/c80004/page={page};price=10000-45000;sell_status=available;state=new;20861=6308,6310/")
    
    time.sleep(2)
    
    soup = BeautifulSoup(driver.page_source, "lxml")
    
    for i in soup.find_all('a', class_="goods-tile__heading ng-star-inserted"):
        link = str(i.get('href'))
        list_urls.append(link)
    


for urls in list_urls:
    driver.get(url=urls)
    time.sleep(2)

    soup_page = BeautifulSoup(driver.page_source, "lxml")

    try:
        bylo_n = soup_page.find("p", class_="product-prices__small").text.replace("₴", "")

        stalo_n = soup_page.find("p", class_="product-prices__big product-prices__big_color_red").text.replace("₴", "")

    except:
        continue

    bylo = ""
    stalo = ""

    for i in stalo_n[1],stalo_n[2], stalo_n[4], stalo_n[5],stalo_n[6] :
        stalo += str(i)

    for d in bylo_n[0], bylo_n[1], bylo_n[3], bylo_n[4], bylo_n[5], bylo_n[6]:
        bylo += str(d)

    skidka = float(bylo) / float(stalo)

    name = soup_page.find("h1", class_="product__title").text

    if skidka > 1.20:
        ws.append([name, skidka, urls])
        wb.save(fn)
        wb.close()


