import os
from collections.abc import Iterable

import pandas as pd
import openpyxl
import openpyxl.styles.numbers
import requests
import time
from bs4 import BeautifulSoup
from PIL import Image
import pillow_avif

# link on web site IEK https://www.iek.ru/products/catalog/search?q=FP-V20-0-10-1-K10

class URLIterator_old:

    def __init__(self, items, base_url):
        self.items = items
        self.base_url = base_url
        self.index = 0

    def __iter__(self):
        return self

    def __next__(self):
        if self.index >= len(self.items):
            raise StopIteration

        item = self.items[self.index]
        url = f"{self.base_url}{item}"
        self.index += 1
        return url, item


class URLIterator:
    def __init__(self, items, base_url):
        """
        :param items: Итерируемый объект (список, генератор и т.д.).
                      Если передаются кортежи (item, code), они будут разложены.
        :param base_url: Базовый URL для генерации ссылок.
        """
        if not isinstance(items, Iterable):
            raise TypeError("items must be iterable (list, generator, etc.)")

        self.items = iter(items)
        self.base_url = base_url

    def __iter__(self):
        return self

    def __next__(self):
        item_data = next(self.items)
        if isinstance(item_data, tuple) and len(item_data) == 2:
            item, code = item_data
        else:
            item, code = item_data, None

        url = f"{self.base_url}{item}"
        return url, item, code

class Files:

    def __init__(self, filepath):
        self.filepath = filepath
        self.status = self.file_exist()

    def file_exist(self):
        if os.path.exists(self.filepath):
            return True
        else:
            return False


class Reader(Files):
    def get_list_excel(self, column_name="Артикул"):
        # check file exist
        try:
            df = pd.read_excel(self.filepath)
            column_values = df[column_name].tolist()
            return column_values
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return []

    def get_list_csv(self, column_art="Артикул", column_code="Код"):
        try:
            df = pd.read_csv(self.filepath, sep=';')
            for _, row in df.iterrows():
                yield (row[column_art], row[column_code])  # Ленивое чтение
        except Exception as e:
            print(f"Error reading CSV: {e}")
            yield from ()  # Пустой генератор


class Excel(Files):
    def __init__(self, filepath):
        super().__init__(filepath)
        if self.status:
            os.remove(self.filepath)
            workbook = openpyxl.Workbook()
            workbook.save(self.filepath)
        else:
            workbook = openpyxl.Workbook()
            workbook.save(self.filepath)

    def clean(self):
        if not self.status:
            pass
        os.remove(self.filepath)

    def list_to_excel(self, value_articul, url=""):
        workbook = openpyxl.load_workbook(self.filepath)
        while len(workbook.sheetnames) > 1:
            workbook.remove(workbook[workbook.sheetnames[1]])

        if not workbook.sheetnames:
            workbook.create_sheet()

        sheet = workbook.active
        first_empty_row = 1
        while sheet.cell(row=first_empty_row, column=1).value is not None:
            first_empty_row += 1
        sheet.cell(row=first_empty_row, column=1).number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[1]
        sheet.cell(row=first_empty_row, column=1, value=value_articul)
        sheet.cell(row=first_empty_row, column=2, value=url)
        workbook.save(self.filepath)
        print(f"Записано в {self.filepath}: {value_articul} {url}")


class Browser:

    @staticmethod
    def get_page(url, timeout=5, try_count=5):
        while try_count > 0:
            response = requests.get(url)
            if response.status_code == 200:
                print(f"Загружена страница: {url}")
                return response.text
            else:
                time.sleep(timeout)
                try_count -= 1
                if try_count == 0:
                    response.raise_for_status()

    @staticmethod
    def download(url, name_file, timeout=5, try_count=5):
        # response = requests.get(url)
        # response.raise_for_status()
        # print(f"Загружен файл с {url} в {name_file}")
        # with open(name_file, 'wb') as file:
        #     file.write(response.content)
        while try_count > 0:
            response = requests.get(url)
            print(response.status_code)
            if response.status_code == 200:
                print(f"Загружен файл с {url} в {name_file}")
                with open(name_file, 'wb') as file:
                    file.write(response.content)
                break
            else:
                time.sleep(timeout)
                try_count -= 1
                if try_count == 0:
                    response.raise_for_status()


class Parser:

    def __init__(self, base_url):
        self.base_url = base_url

    @staticmethod
    def check_element(content_page, type_element, name_class):
        soup = BeautifulSoup(content_page, 'html.parser')
        if soup.find(type_element, class_=name_class):
            return True
        else:
            return False

    @staticmethod
    def get_attr_4el_by_class(page, type_element, class_name, attr_name):
        soup = BeautifulSoup(page, 'html.parser')
        element = soup.select_one(f"{type_element}.{class_name}")
        if element:
            return element.attrs[attr_name]
        else:
            return None

    @staticmethod
    def get_attr_4el_by_id(page, id_name, attr_name):
        soup = BeautifulSoup(page, 'html.parser')
        element = soup.find(id=id_name).select_one("a.popup_link")
        return element.attrs[attr_name]

    @staticmethod
    def get_value_by_class(page, type_element, class_name):
        soup = BeautifulSoup(page, 'html.parser')
        elements = soup.find_all()

    @staticmethod
    def join_base_url(func):
        def wrapper(*args, **kwargs):
            self_instance = args[0]
            original_func = func(*args, **kwargs)
            modif_func = ''.join(f"{self_instance.base_url}{original_func}")
            return modif_func
        return wrapper

    @staticmethod # iek
    def get_link_in_results(search_page, item, tag_name, class_name):
        soup = BeautifulSoup(search_page, 'html.parser')
        spans = soup.find_all(tag_name, class_name)
        for span in spans:
            if item == span.contents[0]:
                parent_a = span.find_parent("a")
                product_url = parent_a.attrs['href']
                return product_url

    @staticmethod # ekf
    def get_links(search_page, item, tag_name1, class_name1, tag_name2, class_name2):
        soup = BeautifulSoup(search_page, 'html.parser')
        p_links = soup.find_all(tag_name1, class_=class_name1)
        span_items = soup.find_all(tag_name2, class_=class_name2)
        for p, span in zip(p_links, span_items):
            tmp_1 = span.contents[1]
            tmp_2 = ''.join(tmp_1.split())
            if item == tmp_2:
                a = p.find("a")
                product_url = a['href']
                return product_url


    def check_element_ref(self, type_element, name_class):
        soup = BeautifulSoup(self.search_page, 'html.parser')
        if soup.find(type_element, class_=name_class):
            return True
        else:
            return False

    def get_element_by_id(self, id):
        soup = BeautifulSoup(self.search_page, 'html.parser')
        element = soup.find(id=id).select_one("a.popup_link")
        return element.attrs['href']

    def get_element(self, type_element, name_class, name_attr):
        soup = BeautifulSoup(self.search_page, 'html.parser')
        element = soup.select_one(f"{type_element}.{name_class}")
        return element.attrs[name_attr]

    def get_link_element(self, type_element, name_class, name_attr='href'):
        element = self.get_element(type_element, name_class, name_attr)
        return f"{self.base_url}{element}"

    def get_element_from(self, page, type_element, name_class, name_attr='href'):
        soup = BeautifulSoup(page, 'html.parser')
        element = soup.select_one(f"{type_element}.{name_class}")
        return element.attrs[name_attr]


class Statistic:

    def __init__(self):
        self.suc = 0
        self.fail_not_found = 0
        self.fail_oth = 0

    def get_stat(self):
        print(
            f"-----------------\nStats: \nSucc: {self.suc} \nFail_4: {self.fail_not_found} \nFail_oth: {self.fail_oth}")

    def add_s(self):
        self.suc += 1

    def add_fn(self):
        self.fail_not_found += 1

    def add_fo(self):
        self.fail_oth += 1


class Converter:

    @staticmethod
    def webp_png(source, destination):
        webp_image = Image.open(source)
        png_image = webp_image.convert("RGBA")
        png_image.save(destination)

    @staticmethod
    def aviv_png(source, destination):
        img = Image.open(source)
        img.save(destination)
