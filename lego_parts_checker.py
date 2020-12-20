# -*- coding: utf-8 -*-

import csv
import datetime
import openpyxl
import pandas as pd
import math
import re
import requests
from bs4 import BeautifulSoup
import sys
import xlwt

class LegoShop:
    def __init__(self, color_dictionary):
        self._colorDictionary = color_dictionary

    def checkStockOf(self, part):
        part_page_url = self.findPartPageUrl(part)
        if (part_page_url is None):
            return Stock(0, 0, None)
        return self.checkStockFrom(part_page_url)

    def name(self):
        return self._name;


class Dgla(LegoShop):
    searchUrl = 'https://dgla.jp/search.cgi?query='
    def __init__(self, color_dictionary):
        super(Dgla, self).__init__(color_dictionary)
        self._name = "dgla"

    def __search(self, part):
        color = self._colorDictionary.convert(part.lddColor(), self._name)
        searchKey = part.partId()
        if (part.partId() == "4697"):
            searchKey = "T型分岐パイプ"
        response = requests.get(Dgla.searchUrl + searchKey + '+' + color)
        return response

    def checkStockFrom(self, part_page_url):
        response = requests.get(part_page_url)
        soup = BeautifulSoup(response.text, 'html.parser')

        quantity_str = soup.find_all('dd', attrs={'class': 'raku-item-vari-stock'})[0].text
        result = re.search(r'\d+', quantity_str.replace(',', ''))
        quantity = int(result.group())

        price_yen = soup.find_all('b', attrs={'class': 'price raku-item-vari-price-num'})[0]
        price = int(price_yen.text.split('円')[0])
        return Stock(quantity, price, part_page_url)

    def findPartPageUrl(self, part):
        color = self._colorDictionary.convert(part.lddColor(), self._name)
        response = self.__search(part)
        soup = BeautifulSoup(response.text, 'html.parser')

        div = soup.find_all('div', attrs={'class': 'results'})[0]
        tr_tags = div.find_all('tr')
        if (len(tr_tags) <= 0):
            return
        min_len = 99999
        tr_tag = ""
        for tr in tr_tags:
            if ('[' + color + ']' in tr.text and len(tr.text) < min_len):
                min_len = len(tr.text)
                tr_tag = tr
        if (tr_tag == ""):
            return

        a = tr_tag.find('a')
        return a['href']


class Brickers(LegoShop):
    siteUrl = 'https://www.brickers.jp/'
    searchUrl = 'https://www.brickers.jp/?mode=srh&keyword='

    def __init__(self, color_dictionary):
        super(Brickers, self).__init__(color_dictionary)
        self._name = "brickers"

    def __findPartPageUrlByPartIdAndColor(self, part):
        color = self._colorDictionary.convert(part.lddColor(), "brickers")
        response = requests.get(Brickers.searchUrl + part.partId() + '-+' + color)
        soup = BeautifulSoup(response.text, 'html.parser')

        ul = soup.find_all('ul', attrs={'class': 'product'})
        if (len(ul) <= 0):
            return
        li = ul[0].find('li')
        a = li.find('a')
        return Brickers.siteUrl + a['href']

    def __findPartPageUrlByBrickId(self, part):
        response = requests.get(Brickers.searchUrl + part.brickId() + "-")
        soup = BeautifulSoup(response.text, 'html.parser')

        ul = soup.find_all('ul', attrs={'class': 'product'})
        if (len(ul) <= 0):
            return
        li = ul[0].find('li')
        a = li.find('a')
        return Brickers.siteUrl + a['href']

    def checkStockFrom(self, part_page_url):
        response = requests.get(part_page_url)
        soup = BeautifulSoup(response.text, 'html.parser')

        div = soup.find('div', attrs={'class': 'spec'})

        td = div.find('td', attrs={'class': 'mark'})
        if (td is None):
            # SOLD OUT
            return Stock(0, 0, part_page_url)
        result = re.search(r'\d+', td.text.replace(',', ''))
        quantity = int(result.group())

        tr = div.find('tr', attrs={'class': 'sales'})
        td = tr.find('td')
        result = re.search(r'\d+', td.text)
        price = int(result.group())

        return Stock(quantity, price, part_page_url)

    def findPartPageUrl(self, part):
        if (part.brickId() is not None):
            url = self.__findPartPageUrlByBrickId(part)
            if (url is not None):
                return url
        return self.__findPartPageUrlByPartIdAndColor(part)


class Part:
    def __init__(self, brick_id, part_id, ldd_color, quantity):
        self.__brickId = brick_id
        self.__partId = part_id
        self.__lddColor = ldd_color
        self.__quantity = quantity
        self.__buy_quantity = 0

    def brickId(self):
        return self.__brickId

    def buy(self, quantity):
        self.__buy_quantity += quantity

    def lddColor(self):
        return self.__lddColor

    def lack(self):
        return max([self.__quantity - self.__buy_quantity, 0])

    def partIdLog(self):
        if (self.__partId == self.partId()):
            return self.__partId
        return self.__partId + " -> " + self.partId()

    def partId(self):
        if (self.__partId == "50746"):
            # LegoSlope 30 1 x 1 x 2/3
            return "54200"
        elif (self.__partId == "3070"):
            # LegoTile 1 x 1 with Groove (3070)
            return "3070b"
        elif (self.__partId == "11153"):
            # LegoSlope, Curved 4 x 1
            return "61678"
        elif (self.__partId == "58856"):
            # LegoBar 4L (Lightsaber Blade / Wand)
            return "30374"
        elif (self.__partId == "3794"):
            # LegoPlate, Modified 1 x 2 with 1 Stud with Groove and Bottom Stud Holder (Jumper)
            return "15573"
        elif (self.__partId == "60897"):
            # Plate, Modified 1 x 1 with Clip Vertical (Undetermined Clip Type)
            return "4085"
        elif (self.__partId == "6141"):
            # Plate, Round 1 x 1
            return "4073"
        elif (self.__partId == "30359"):
            # LegoBar 1 x 8 with Brick 1 x 2 Curved Top End (Undetermined Type)
            return "30359b"
        return self.__partId

    def quantity(self):
        return self.__quantity


class PartsService:
    def readBrickId(part_data):
        if (math.isnan(float(part_data[1]))):
            return
        return str(int(part_data.Brick))

    def readLddColor(part_data):
        return part_data[5]

    def readPartId(part_data):
        if (math.isnan(part_data.Part)):
            return
        return str(int(part_data.Part))

    def readQuantity(part_data):
        return int(part_data.Quantity)

    def createPart(part_data):
        if (PartsService.readPartId(part_data) is None):
            return
        return Part(\
                PartsService.readBrickId(part_data),\
                PartsService.readPartId(part_data),\
                PartsService.readLddColor(part_data),\
                PartsService.readQuantity(part_data))


class Stock:
    def __init__(self, quantity, price, url):
        self.__quantity = quantity
        self.__price = price
        self.__url = url

    def quantity(self):
        return self.__quantity

    def price(self):
        return self.__price

    def url(self):
        return self.__url


class ColorDictionary:
    def __init__(self, source):
        self.__df = pd.read_csv(source, index_col="ldd", skipinitialspace=True)

    def convert(self, ldd_color, web_site):
        return self.__df.loc[ldd_color, web_site]


class PartsCheckReport:
    def __init__(self, shop_list):
        cols = ["partId", "brickId", "color", "quantity(request)"]
        for shop in shop_list:
            cols += [\
                    "url(" + shop.name() + ")",\
                    "quantity(" + shop.name() + ")",\
                    "price(" + shop.name() + ")",\
                    "buy_quantity(" + shop.name() + ")",\
                    "buy_price(" + shop.name() + ")"]
        cols.append("lack")
        self.__df = pd.DataFrame(index=[], columns=cols)

    def append(self, row):
        record = pd.Series(row, index=self.__df.columns)
        self.__df = self.__df.append(record, ignore_index=True)

    def output(self, file_name):
        self.__df.to_excel(file_name)


class PartsStockCheckService:
    def __init__(self):
        color_dictionary = ColorDictionary('color.csv')
        self.__lego_shops = [\
                Dgla(color_dictionary),\
                Brickers(color_dictionary)]

    def __checkShopStockFor(self, part):
        result = []
        for shop in self.__lego_shops:
            result += self.__checkShopStockForEach(shop, part)
        return result

    def __checkShopStockForEach(self, shop, part):
            print("  shop: " + shop.name())
            stock = shop.checkStockOf(part)
            buy_quantity = 0
            buy_price = 0
            if (part.lack() > 0 and stock.quantity() > 0):
                buy_quantity = min(part.lack(), stock.quantity())
                part.buy(buy_quantity)
                buy_price = buy_quantity * stock.price()
            return [\
                    stock.url(),\
                    stock.quantity(),\
                    stock.price(),\
                    buy_quantity,\
                    buy_price]

    def __checkStockForEach(self, part):
        print("partId: " + part.partId() + " (color: " + part.lddColor() + ")")
        print("  date: " + str(datetime.datetime.now()))
        return [\
                part.partIdLog(),\
                part.brickId(),\
                part.lddColor(),\
                part.quantity()]\
                + self.__checkShopStockFor(part)\
                + [part.lack()]

    def checkStockFor(self, parts_list):
        report = PartsCheckReport(self.__lego_shops)
        for part_data in parts_list.itertuples():
            part = PartsService.createPart(part_data)
            if (part is None):
                continue
            result = self.__checkStockForEach(part)
            report.append(result)
        return report


class LegoPartsCheckerApplication:
    def __load(self, bom_file_name):
        SHEET_NAME = 'Sheet1'
        return pd.read_excel(bom_file_name, sheet_name=SHEET_NAME)

    def check(self, bom_file_name):
        parts_list = self.__load(bom_file_name)
        service = PartsStockCheckService()
        report = service.checkStockFor(parts_list)
        report.output("check_" + bom_file_name)


def main():
    if len(sys.argv) == 1:
        print("usage:")
        print("      $ python3 lego_parts_checker.py BOM_FILE")
        sys.exit()

    bom_file_name = sys.argv[1]
    app = LegoPartsCheckerApplication()
    app.check(bom_file_name)

main()
