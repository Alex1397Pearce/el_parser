import os

import requests
from bs4 import BeautifulSoup
import re
from main import URLIterator, Parser, Browser, Reader, Excel, Statistic, Converter


file = Reader(r"\\1csrv\SystemFiles\pictures\data\iek.csv")
success_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Succes_iek.xlsx")
failed_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Failed_iek.xlsx")
stat = Statistic()


# def get_product_link_iek(item):
#     soup = BeautifulSoup(search_page, 'html.parser')
#     spans = soup.find_all("span", "ProductArticle_btn-text__oYFaw")
#     for span in spans:
#         if item == span.contents[0]:
#             parent_a = span.find_parent("a")
#             product_url = parent_a.attrs['href']
#             return product_url


# Тестовые данные
def my_generator():
    yield ("WYP10-10-03-03-Z-G", "УТ-0144254")
    yield ("LB-1000A5-25-F-LUF", "УТ-0141224")
# data = my_generator()

# Рабочие данные
data = file.get_list_csv()

iterator = URLIterator(data, "https://www.iek.ru/products/catalog/search?q=")
for url, item, code in iterator:
    try:
        search_page = Browser.get_page(url)
        if not Parser.check_element(search_page, 'div', 'NothingFound_message__2aExd'):
            product_url = Parser.get_link_in_results(search_page, item, "span", "ProductArticle_btn-text__oYFaw")
            if product_url:
                product_page = Browser.get_page(f"https://iek.ru{product_url}")
                soup = BeautifulSoup(product_page, 'html.parser')
                if Parser.check_element(product_page, 'div', 'MirgationProduct_container__5HpZ3'):
                    product_page = Browser.get_page(f"https://generica.su/products/catalog/article/{item}")
                    soup = BeautifulSoup(product_page, 'html.parser')
                    div = soup.find('div', "ProductMedia_main-photo__gZY6E")
                    img = div.find("img").next_sibling
                    if img:
                        image = img.attrs['srcset']
                        image_url = image.split()
                        image_path = fr"\\1csrv\SystemFiles\pictures\{code}.avif"
                        Browser.download(image_url[0], image_path)
                        Converter.convert_to_jpg(image_path)
                        success_file.list_to_excel(code, image_path)
                        stat.add_s()
                        stat.get_stat()

                    else:
                        print("Нет картинки на сайте")
                        failed_file.list_to_excel(item, "Нет картинки на сайте")
                        stat.add_fo()
                        stat.get_stat()
                else:
                    div = soup.find('div', "ProductMedia_main-photo__gZY6E")
                    img = div.find("img").next_sibling
                    if img:
                        image = img.attrs['srcset']
                        image_url = image.split()
                        image_path = fr"\\1csrv\SystemFiles\pictures\{code}.avif"
                        Browser.download(image_url[0], image_path)
                        Converter.convert_to_jpg(image_path)
                        success_file.list_to_excel(code, image_path)
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
    except (requests.exceptions.HTTPError,requests.exceptions.ConnectTimeout) as e:
        print(e)
        failed_file.list_to_excel(item, url)
        stat.add_fo()
        stat.get_stat()