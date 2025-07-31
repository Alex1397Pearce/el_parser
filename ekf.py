import os

import requests
from bs4 import BeautifulSoup
import re
from main import URLIterator, Parser, Browser, Reader, Excel, Statistic, Converter


file = Reader(r"\\1csrv\SystemFiles\pictures\data\ekf.csv")
success_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Succes_ekf.xlsx")
failed_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Failed_ekf.xlsx")
stat = Statistic()


def get_product_link_iek(item):
    soup = BeautifulSoup(search_page, 'html.parser')
    spans = soup.find_all("span", "ProductArticle_btn-text__oYFaw")
    for span in spans:
        if item == span.contents[0]:
            parent_a = span.find_parent("a")
            product_url = parent_a.attrs['href']
            return product_url



# Тестовые данные
def my_generator():
    # yield ("plc-jxb-4/35RD-gy")
    yield ("AleSta-PM19iU47", "УТ-0164323")
    # yield ("Б0052635хуета", "УТ-000000")
# data = my_generator()

# Рабочие данные
data = file.get_list_csv()

iterator = URLIterator(data, "https://ekfgroup.com/ru/search?q=")
for url, item, code in iterator:
    try:
        search_page = Browser.get_page(url)
        product_url = Parser.get_links(search_page, item, 'p', 'product-title', 'span', 'product-vendor-code')
        if product_url:
            product_page = Browser.get_page(f"https://ekfgroup.com{product_url}")
            image_url = Parser.get_attr_4el_by_class(product_page, 'img', 'gallery-slide-image', 'src')
            if image_url is not None:
                image_path = fr"\\1csrv\SystemFiles\pictures\{code}.png"
                Browser.download(image_url, image_path)
                Converter.convert_to_jpg(image_path)
                success_file.list_to_excel(code, image_path)
                stat.add_s()
                stat.get_stat()
            else:
                print("Проблемная картинка на сайте")
                failed_file.list_to_excel(item, "Проблемная картинка на сайте")
                stat.add_fo()
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