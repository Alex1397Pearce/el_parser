import requests
from main import URLIterator, Parser, Browser, Reader, Excel


file = Reader(r"c:\Users\pr54m\Desktop\inhome.xlsx")
success_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Succes_InHome.xlsx")
failed_file = Excel(r"\\1csrv\SystemFiles\pictures\results\Failed_InHome.xlsx")


data = file.get_list()
iterator = URLIterator(data, "https://in-home.ru/products/?q=")
p = Parser("https://in-home.ru")
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
        else:
            failed_file.list_to_excel(item)
    except requests.exceptions.HTTPError as e:
        print(e)
        failed_file.list_to_excel(item, url)