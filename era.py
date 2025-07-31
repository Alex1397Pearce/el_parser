import os

import requests
from bs4 import BeautifulSoup
import re
from main import URLIterator, Parser, Browser, Reader, Excel, Statistic, Converter


file = Reader(r"\\1csrv\SystemFiles\pictures\data\era.csv")
success_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Succes_era.xlsx")
failed_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Failed_era.xlsx")
stat = Statistic()

data = file.get_list_csv()

def my_generator():
    yield ("Б0052635хуета", "УТ-0144254")
    yield ("Б0052635", "УТ-0141224")
# data = my_generator()

iterator = URLIterator(data, "https://www.eraworld.ru/search?q=")
p = Parser("https://www.eraworld.ru")
for url, item, code in iterator:
    try:
        search_page = Browser.get_page(url)
        product_url = p.get_attr_4el_by_class(search_page, 'div', 'media-left a', 'href')
        if product_url:
            product_page = Browser.get_page(product_url)
            image_url = p.get_attr_4el_by_class(product_page, 'div', 'big_image a', 'href')
            image_path = fr"\\1csrv\SystemFiles\pictures\{code}.png"
            Browser.download(image_url, image_path)
            Converter.convert_to_jpg(image_path)
            success_file.list_to_excel(code, image_path)
            stat.add_s()
            stat.get_stat()
        else:
            print("Нет нужной позиции на сайте")
            failed_file.list_to_excel(item, "Нет нужной позиции на сайте")
            stat.add_fo()
            stat.get_stat()
    except (requests.exceptions.HTTPError,requests.exceptions.ConnectTimeout) as e:
        print(e)
        failed_file.list_to_excel(item, url)
        stat.add_fo()
        stat.get_stat()