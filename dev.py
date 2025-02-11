from main import URLIterator, Parser, Browser


def set_iter():
    # file = Reader(r"c:\Users\pr54m\Desktop\inhome.xlsx")
    # data = file.get_list()

    return iterator


data = ("4690612024004", "4690612024028")
iterator = URLIterator(data, "https://in-home.ru/products/?q=")
p = Parser("https://in-home.ru")
for url, item in iterator:
    # p.make_request(url)
    search_page = Browser.get_page(url)
    if not Parser.check_element(search_page, 'div', 'no_goods'):
        pass

    # if not p.check_element_ref(type_element='div', name_class='no_goods'):
    #     product_link = p.get_link_element(type_element='div', name_class='catalog-block__info-title a')
    #     p.make_request(product_link)
    #     image_link = p.get_element_by_id('big-photo-0')
    #     print(image_link)
