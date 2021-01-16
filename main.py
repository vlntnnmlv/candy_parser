import re
import requests
import xlsxwriter
from bs4 import BeautifulSoup as bs

data = [
	"https://vtk-moscow.ru/shop/pishhevye-ingredienty/slivochnyj-syr-i-slivki/syr-tvorozhnyj-chudskoe-ozero-60-33-kg/",
	"https://bakerstore.ru/ingredienty/pektin/pektin-yablochnyj-100-g",
	"https://msk.tortomaster.ru/catalog/kakao-maslo/",
	"https://bakerstore.ru/ingredienty/suhie-kremy/krem-dlya-torta-so-vkusom-vanili-50-g-dr-oetker",
	"https://bakerstore.ru/podstavki-dlya-tortov/podstavka-vrashchayushchayasya-dlya-dekorirovaniya-tortov-28sm"
]

def matches(s):
	pattern1 = re.compile("vtk")
	pattern2 = re.compile("bakerstore")
	pattern3 = re.compile("tortomaster")
	if (pattern1.search(s)):
		return (0);
	if (pattern2.search(s)):
		return (1);
	if (pattern3.search(s)):
		return (2);

def get_data(page):
	headers = {"User-agent":"Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.80 Safari/537.36"}
	res = requests.get(page, headers=headers)

	txt = open("txt","w")
	txt.write(res.text)
	txt.close

	with open("txt", "r") as f:
		contents = f.read()
	soup = bs(contents, "html.parser")
	return (soup)

def iget_bakerstore(parser):
	raw = parser.find("span", {"class" : "autocalc-product-special"})
	if (raw):
		price = raw.string
	else:
		price = parser.find("span", {"class" : "autocalc-product-price"}).string
	return (price)

def iget_vtk(parser):
	price = parser.find("span", {"class" : "tprice-value"}).string
	return (price.replace(" ", ""))

def iget_tortomaster(parser):
	price = parser.find_all("span", {"class" : "price"})[0].text.split(' ')[0]
	return (price)

output_data = []

iget = [iget_vtk, iget_bakerstore, iget_tortomaster];

if __name__ == "__main__":
	for page in data:
		i = matches(page)
		price = iget[i](get_data(page))
		output_data.append((page, price))

	workbook = xlsxwriter.Workbook('out.xlsx')
	worksheet = workbook.add_worksheet()
	money_format = workbook.add_format({'num_format': '#,##0'})
	row = 0
	col = 0
	for item, cost in (output_data):
		worksheet.write(row, col,     item.rstrip('/').split('/')[-1])
		worksheet.write_number(row, col + 1, float(cost), money_format)
		row += 1

	workbook.close()