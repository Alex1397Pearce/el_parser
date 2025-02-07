from main import URLIterator, Reader, Parser


def set_iter():
    # file = Reader(r"c:\Users\pr54m\Desktop\inhome.xlsx")
    # data = file.get_list()
    data = ("4690612024004", "4690612024028")
    iterator = URLIterator(data, "https://in-home.ru/products/?q=")
    return iterator


iterator = set_iter()
r = Parser("https://in-home.ru")
for url, item in iterator:
    r.make_request(url)
    if not r.check_element(type_element='div', name_class='no_goods'):
        product_link = r.get_link_element(type_element='div', name_class='catalog-block__info-title a')
        r.make_request(product_link)
        image_link = r.get_element_by_id('big-photo-0')
        print(image_link)
