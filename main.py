#!/usr/bin/env python3

import	pandas
import	requests
import	warnings
import	openpyxl
from 	openpyxl.styles import Font
from	openpyxl.utils.dataframe import dataframe_to_rows
import	os
from	bs4 import BeautifulSoup as bs
from	itertools import islice
from	tkinter import *
from	tkinter import filedialog
import	sys
import	tkinter.ttk as ttk
import	threading
import	time
from	concurrent import futures

def get_raw_data(path):
	''' Reads data from xlsx file to pandas DataFrame '''
	wb = openpyxl.load_workbook(path)
	ws = wb.active

	data = ws.values
	cols = next(data)[1:]
	data = list(data)
	idx = [r[0] for r in data]
	data = (islice(r, 1, None) for r in data)
	df = pandas.DataFrame(data, index=idx, columns=cols)
	return (df)

def get_data(page):
	''' Downloads a page data from the website 'page'
		and returns a parser object '''
	if (not page):
		return (None)
	if (page[:8] != "https://"):
		return (page)
	headers = {"User-agent":"Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.80 Safari/537.36"}
	res = requests.get(page, headers=headers, verify=False)
	soup = bs(res.text, "html.parser")
	return (soup)

def iget_bakerstore(parser):
	''' Gets price from the bakerstore page's parser object '''
	if (not parser):
		return ("")
	if (type(parser) == str):
		return (parser)
	raw = parser.find("span", {"class" : "autocalc-product-special"})
	if (raw):
		price = raw.string
	else:
		price = parser.find("span", {"class" : "autocalc-product-price"}).string
	return int(price.replace(".0", ""))

def iget_vtk(parser):
	''' Gets price from the vtk page's parser object '''
	if (not parser):
		return ("")
	if (type(parser) == str):
		return (parser)
	price = parser.find("span", {"class" : "tprice-value"}).string
	return int(price.replace(" ", ""))

def iget_tortomaster(parser):
	''' Gets price from the tortomaster page's parser object '''
	if (not parser):
		return ("")
	if (type(parser) == str):
		return (parser)
	price = "".join(parser.find_all("span", {"class" : "price"})[0].text[:-1].split(" "))
	return int(price)

def onclick(event=None):
	''' Checks is chosen file valid and calls parser function '''
	try:
		res = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
		if (res):
			exec_path = res
			path.config(text = "PATH: " + os.path.abspath(res))
			if (res.split(".")[-1] != "xlsx"):
				lbl.config(text = "Введите название Ексель файла.", fg="red")
			elif not os.path.exists(res):
				lbl.config(text ="Такого Ексель файла не сущесвует!", fg="red")
			else:
				btn.config(state=DISABLED)
				lbl.config(text = "Загружаем...", fg="blue")
				do_job(res)
				btn.config(state=NORMAL)
	except BaseException as e:
		lbl.config(text = "Что-то пошло не так!\nОбратитесь к разработчику!\n" + str(e), fg = "red")

def do_job(filename):
	''' Parser function. Sends requests by link in the file, 
		and changes data with price values '''
	pb['value'] = 0
	raw_data = get_raw_data(filename)
	new_data = pandas.DataFrame()
	new_data["VTK"]			= raw_data["VTK"].apply(lambda x: iget_vtk(parser = get_data(x)))
	pb['value'] += 10
	new_data["bakerstore"]	= raw_data["bakerstore"].apply(lambda x: iget_bakerstore(parser = get_data(x)))
	pb['value'] += 10
	new_data["tortomaster"]	= raw_data["tortomaster"].apply(lambda x: iget_tortomaster(parser = get_data(x)))
	pb['value'] += 10

	empty = pandas.DataFrame()
	empty[""] = ""
	new_data = new_data.join(empty[""])
	new_data = new_data.join(raw_data["VTK"], rsuffix="_link")
	new_data = new_data.join(raw_data["bakerstore"], rsuffix="_link")
	new_data = new_data.join(raw_data["tortomaster"], rsuffix="_link")
	wb = openpyxl.Workbook()
	pb['value'] += 10
	ws = wb.active
	pb['value'] += 10
	for r in dataframe_to_rows(new_data, index=True, header=True):
		ws.append(r)
	pb['value'] += 10
	for cell in ws['A'] + ws[1]:
		if (cell.value):
			cell.style = 'Pandas'

	for cell in ws['B'] + ws['C'] + ws['D']:
		cell.number_format = "General"

	font = Font(name='Calibri',
                 size=11,
                 bold=False,
                 italic=False,
                 vertAlign=None,
                 underline='none',
                 strike=False,
                 color='408B22')

	for row in range(3, new_data.shape[0] + 3):
		m = 1000000000
		min_col = 2
		for col in [2,3,4]:
			if ws.cell(row = row, column = col).value and \
				type(ws.cell(row = row, column = col).value) == int and \
				int(ws.cell(row = row, column = col).value) < int(m):
				m = int(ws.cell(row = row, column = col).value)
				min_col = col
		ws.cell(row = row, column = min_col).font = font

	ws.column_dimensions['A'].width = 100
			
	pb['value'] += 10
	wb.save("./" + filename.split('/')[-1].split('.')[-2] + "_out.xlsx")
	wb.close()
	pb['value'] += 10
	lbl.config(text ="Готово!", fg="green", font=("bold"))
	pb['value'] += 10

if __name__ == "__main__":	
	os.chdir(os.getcwd())
	warnings.filterwarnings("ignore")

	window = Tk()
	window.title("Candy Parser")
	window.geometry('480x320')
	back=Frame(master=window, width=500, height=500)
	back.pack(fill="none", expand=TRUE)
	lbl = Label(text = "Привет! Нажми на кнопку, чтобы загрузить данные!")
	lbl.pack()
	path = Label(text = "PATH: ", fg="grey")
	path.pack()
	pb = ttk.Progressbar(back, mode="determinate", maximum = 90, cursor = "exchange")
	pb.pack(pady=20)
	btn = Button(back, text="Выбрать файл", command=lambda : threading.Thread(target=onclick).start())
	btn.pack()
	window.bind('<Escape>', sys.exit)
	window.mainloop()
