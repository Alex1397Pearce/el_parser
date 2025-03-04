import os

import requests
from bs4 import BeautifulSoup
import re
from main import URLIterator, Parser, Browser, Reader, Excel, Statistic, Converter


file = Reader(r"c:\Users\pr54m\Desktop\iek.xlsx")
success_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Succes_iek.xlsx")
failed_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Failed_iek.xlsx")
stat = Statistic()


def get_product_link_iek(item):
    soup = BeautifulSoup(search_page, 'html.parser')
    spans = soup.find_all("span", "ProductArticle_btn-text__oYFaw")
    for span in spans:
        if item == span.contents[0]:
            parent_a = span.find_parent("a")
            product_url = parent_a.attrs['href']
            return product_url


data = file.get_list()
# data = ("MVA20-4-013-B", "AR-M10N-MA-2-D016-hh")
iterator = URLIterator(data, "https://www.iek.ru/products/catalog/search?q=")
p = Parser("https://www.iek.ru")
for url, item in iterator:
    try:
        search_page = Browser.get_page(url)
        if not Parser.check_element(search_page, 'div', 'NothingFound_message__2aExd'):
            product_url = get_product_link_iek(item)
            if product_url:
                product_page = Browser.get_page(f"https://iek.ru{product_url}")
                soup = BeautifulSoup(product_page, 'html.parser')
                div = soup.find('div', "ProductMedia_main-photo__gZY6E")
                img = div.find("img").next_sibling
                if img:
                    image = img.attrs['srcset']
                    image_url = image.split()
                    image_path_temp = fr"\\1csrv\SystemFiles\pictures\{item}.avif"
                    image_path_final = fr"\\1csrv\SystemFiles\pictures\{item}.png"

                    Browser.download(image_url[0], image_path_temp)
                    Converter.aviv_png(image_path_temp, image_path_final)
                    os.remove(image_path_temp)
                    print("Картинка удалена")
                    success_file.list_to_excel(item, image_path_final)
                    stat.add_s()
                    stat.get_stat()

                else:
                    print("Нет картинки на сайте")
                    failed_file.list_to_excel(item, "Нет картинки на сайте")
                    stat.add_fo()
                    stat.get_stat()
            else:
                print("Нет нужной позиции на сайте")
                failed_file.list_to_excel(item, "Нет нужной позиции на сайте")
                stat.add_fo()
                stat.get_stat()
        else:
            failed_file.list_to_excel(item)
            stat.add_fn()
            stat.get_stat()
    except requests.exceptions.HTTPError as e:
        print(e)
        failed_file.list_to_excel(item, url)
        stat.add_fo()
        stat.get_stat()