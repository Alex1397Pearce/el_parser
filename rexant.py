import os

import requests
from bs4 import BeautifulSoup
import re
from main import URLIterator, Parser, Browser, Reader, Excel, Statistic, Converter


file = Reader(r"\\1csrv\SystemFiles\pictures\data\rexant.csv")
success_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Succes_rexant.xlsx")
failed_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Failed_rexant.xlsx")
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
    yield ("12-0620", "УТ-0119654") # Не найдено
    yield ("51-1000", "УТ-0114087") # Картинка есть
    yield ("13-1347", "УТ-0108187") # найдено, но картинки нет

# data = my_generator()

# Рабочие данные
data = file.get_list_csv()

iterator = URLIterator(data, "https://rexant.ru/catalog/?q=")
# p = Parser("https://ekfgroup.com/")
for url, item, code in iterator:
    try:
        search_page = Browser.get_page(url)
        # Если найдена фраза "Сожалеем, но ничего не найдено.", то надо искать следующую
        # product_url = Parser.get_attr_4el_by_class(search_page, 'p', 'product-title a', 'href')
        product_url = Parser.get_attr_4el_by_class(search_page, 'a', 'product-item-image-wrapper', 'href')
        if product_url:
            product_page = Browser.get_page(f"https://rexant.ru{product_url}")
            image_url = Parser.get_child_img(product_page, 'product-item-detail-slider-image')
            # image_url2 = Parser.get_attr_4el_by_class(product_page, 'img', 'spinner-img', 'src')
            if image_url != '/upload/webp/local/templates/rexant/components/bitrix/catalog.element/bootstrap_v4/images/no_photo.webp':
                image_path = fr"\\1csrv\SystemFiles\pictures\test\{code}.png"
                Browser.download(f"https://rexant.ru{image_url}", image_path)
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