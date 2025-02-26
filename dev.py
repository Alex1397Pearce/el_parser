import requests
from bs4 import BeautifulSoup
import re
from main import URLIterator, Parser, Browser, Reader, Excel, Statistic


file = Reader(r"c:\Users\pr54m\Desktop\iek.xlsx")
success_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Succes_iek.xlsx")
failed_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Failed_iek.xlsx")
stat = Statistic()

# data = file.get_list()
data = ("TI5-50-V-024-31-7035", "TI5-12-N-030-030-015-66-hueta", "FP-V20-0-10-1-K10")
iterator = URLIterator(data, "https://www.iek.ru/products/catalog/search?q=")
p = Parser("https://www.iek.ru")
for url, item in iterator:
    try:
        # поиск страницы из поиска
        search_page = Browser.get_page(url)
        if not Parser.check_element(search_page, 'div', 'NothingFound_message__2aExd'):
            pass
        else:
            pass

        soup = BeautifulSoup(search_page, 'html.parser')
        spans = soup.find_all("span", "ProductArticle_btn-text__oYFaw")
        for span in spans:
            text = span.contents[0]
            if item == span.contents[0]:
                print("EBABULA")
                parent_a = span.find_parent("a")
                product_url = parent_a.attrs['href']
                print(product_url)

        # поиск картинки на странице
        product_page = Browser.get_page(f"https://iek.ru{product_url}")
        soup = BeautifulSoup(product_page, 'html.parser')
        div = soup.find('div', "ProductMedia_main-photo__gZY6E")
        img = div.find("img").next_sibling
        img_link = img.attrs['srcset']
        print("" + img_link)
    except requests.exceptions.HTTPError as e:
        print(e)