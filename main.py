import requests
from bs4 import BeautifulSoup
import openpyxl


urls = ['https://nhatrangshop.ru/category/zdorove/fito-pharma/', 'https://nhatrangshop.ru/category/zdorove/artishok/',
       'https://nhatrangshop.ru/category/zdorove/grib-linchzhi/', 'https://nhatrangshop.ru/category/zdorove/volosy-kozha-i-nogti/',
       'https://nhatrangshop.ru/category/zdorove/glaza-i-zrenie/', 'https://nhatrangshop.ru/category/zdorove/davlenie/',
       'https://nhatrangshop.ru/category/zdorove/detoksikatsiya-i-ochishchenie/','https://nhatrangshop.ru/category/zdorove/deyatelnost-mozga/',
       'https://nhatrangshop.ru/category/zdorove/imunnaya-sistema/', 'https://nhatrangshop.ru/category/zdorove/kishechnik/',
       'https://nhatrangshop.ru/category/zdorove/sustavy/', 'https://nhatrangshop.ru/category/zdorove/mochevoy-puzyr/',
       'https://nhatrangshop.ru/category/zdorove/muzhskoe-i-zhenskoe-zdorove/', 'https://nhatrangshop.ru/category/zdorove/organy-dykhaniya/',
       'https://nhatrangshop.ru/category/zdorove/polost-rta/', 'https://nhatrangshop.ru/category/zdorove/prostuda-i-gripp/',
       'https://nhatrangshop.ru/category/zdorove/sakharnyy-diabet/', 'https://nhatrangshop.ru/category/zdorove/serdtse/',
       'https://nhatrangshop.ru/category/zdorove/sistema-pishchevareniya/', 'https://nhatrangshop.ru/category/zdorove/snizhenie-vesa/',
       'https://nhatrangshop.ru/category/zdorove/son/', 'https://nhatrangshop.ru/category/bad-from-vietnam/category_228/',
       'https://nhatrangshop.ru/category/bad-from-vietnam/category_226/', 'https://nhatrangshop.ru/category/bad-from-vietnam/komplekty/',
       'https://nhatrangshop.ru/category/bad-from-vietnam/category_216/', 'https://nhatrangshop.ru/category/category_212/category_214/',
       'https://nhatrangshop.ru/category/category_212/yeda_1/', 'https://nhatrangshop.ru/category/category_212/category_215/',
       'https://nhatrangshop.ru/category/category_212/category_213/', 'https://nhatrangshop.ru/category/beauty/category_241/',
       'https://nhatrangshop.ru/category/beauty/category_232/', 'https://nhatrangshop.ru/category/beauty/category_230/',
       'https://nhatrangshop.ru/category/beauty/category_231/', 'https://nhatrangshop.ru/category/beauty/category_233/',
        'https://nhatrangshop.ru/category/kosmetika-iz-korei/', 'https://nhatrangshop.ru/category/kozhanye-izdeliya/']
wb = openpyxl.Workbook()
sheet = wb.active

sheet["A1"] = "Название"
sheet["B1"] = "Описание"
sheet["C1"] = "Цена"
sheet["D1"] = "Ссылка"

k = 2
for url in urls:

    html = requests.get(url).text
    soup = BeautifulSoup(html, 'html.parser')

    for i, item in enumerate(soup.select(".product-tile__outer"),3):
      name = item.select_one(".product-tile__name").text
      desc = item.select_one(".product-tile__description").text
      price = item.select_one(".price").text
      badges_div = item.find('div', class_='product-tile__image js-tile-gallery-block')
      link = badges_div.find('a')
      link1 = 'https://nhatrangshop.ru' + link['href']
      sheet[f"A{k}"] = name
      sheet[f"B{k}"] = desc
      sheet[f"C{k}"] = price
      sheet[f"D{k}"] = link1
      k += 1


wb.save("products.xlsx")