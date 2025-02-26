import requests
from main import URLIterator, Parser, Browser, Reader, Excel, Statistic

file = Reader(r"c:\Users\pr54m\Desktop\iek.xlsx")
success_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Succes_iek.xlsx")
failed_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Failed_iek.xlsx")
stat = Statistic()

data = file.get_list()
# data = ("4690612024004", "4690612024028")
iterator = URLIterator(data, "https://www.iek.ru/products/catalog/search?q=")
p = Parser("https://www.iek.ru")
for url, item in iterator:
    try:
        search_page = Browser.get_page(url)
        if not Parser.check_element(search_page, 'div', 'no_goods') and not Parser.check_element(search_page, 'div',
                                                                                                 'alert-danger'):
            product_url = p.get_attr_4el_by_class(search_page, 'div', 'catalog-block__info-title a', 'href')
            product_page = Browser.get_page(f"https://in-home.ru{product_url}")
            image_url = p.get_attr_4el_by_id(product_page, 'big-photo-0', 'href')
            image_path = fr"\\1csrv\SystemFiles\pictures\{item}.png"
            Browser.download(image_url, image_path)
            success_file.list_to_excel(item, image_path)
            stat.add_s()
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

