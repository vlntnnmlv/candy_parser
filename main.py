#!/usr/bin/env python3

import	pandas
import	requests
import	openpyxl
from	openpyxl.utils.dataframe import dataframe_to_rows
import	os
from	bs4 import BeautifulSoup as bs
from	itertools import islice
from	tkinter import *
import	sys
import	tkinter.ttk as ttk
import	threading

def get_raw_data(path):
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
	''' Downloads a page data from the website 'page' '''
	''' and returns a parser object '''
	headers = {"User-agent":"Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.80 Safari/537.36"}
	res = requests.get(page, headers=headers)

	txt = open("txt","w")
	txt.write(res.text)
	txt.close

	with open("txt", "r") as f:
		contents = f.read()
	soup = bs(contents, "html.parser")
	os.remove("txt")
	return (soup)

def iget_bakerstore(parser):
	''' Gets price from the bakerstore page's parser object '''
	raw = parser.find("span", {"class" : "autocalc-product-special"})
	if (raw):
		price = raw.string
	else:
		price = parser.find("span", {"class" : "autocalc-product-price"}).string
	return (price + ".0")

def iget_vtk(parser):
	''' Gets price from the vtk page's parser object '''
	price = parser.find("span", {"class" : "tprice-value"}).string
	return (price.replace(" ", "") + ".0")

def iget_tortomaster(parser):
	''' Gets price from the tortomaster page's parser object '''
	price = "".join(parser.find_all("span", {"class" : "price"})[0].text[:-1].split(" "))
	return (price)

def onclick(event=None):
	res = txt.get()
	exec_path = os.path.abspath(os.path.dirname(sys.argv[0]))
	path.config(text = "PATH: " + exec_path + "/" + res)
	if (res.split(".")[-1] != "xlsx"):
		lbl.config(text = "Введите название Ексель файла.", fg="red")
	elif not os.path.exists(exec_path + "/" + res):
		lbl.config(text ="Такого Ексель файла не сущесвует!", fg="red")
	else:
		btn.config(state=DISABLED)
		lbl.config(text = "Загружаем...", fg="blue")
		threading.Thread(target=do_job(exec_path, res), daemon=True).start()
		lbl.config(text ="Готово!", fg="green")
	
def do_job(dirpath, filename):
	raw_data = get_raw_data(dirpath + "/" + filename)
	pb['value'] += 4
	raw_data["VTK"]			= raw_data["VTK"].apply(lambda x: iget_vtk(parser = get_data(x)))
	pb['value'] += 4
	raw_data["bakerstore"]	= raw_data["bakerstore"].apply(lambda x: iget_bakerstore(parser = get_data(x)))
	pb['value'] += 4
	raw_data["tortomaster"]	= raw_data["tortomaster"].apply(lambda x: iget_tortomaster(parser = get_data(x)))

	wb = openpyxl.Workbook()
	ws = wb.active
	
	for r in dataframe_to_rows(raw_data, index=True, header=True):
		ws.append(r)
		pb['value'] += 4

	for cell in ws['A'] + ws[1]:
		cell.style = 'Pandas'
		pb['value'] += 4

	wb.save(dirpath + "/" + "out.xlsx")

if __name__ == "__main__":
	os.chdir(os.getcwd())

	window = Tk()
	window.title("Candy Parser")
	window.geometry('640x480')

	back=Frame(master=window, width=500, height=500)
	back.pack(fill="none", expand=TRUE)
	
	lbl = Label(text = "Привет! Нажми на кнопку, чтобы загрузить данные!")
	lbl.pack()
	path = Label(text = "PATH: ", fg="grey")
	path.pack()
	txt = Entry(back,width=20)
	txt.pack()
	txt.focus()
	pb = ttk.Progressbar(back, mode="determinate")
	pb.pack()
	btn = Button(back, text="Загрузить", command=onclick)
	window.bind('<Return>', onclick)
	window.bind('<Escape>', sys.exit)
	btn.pack()
	window.mainloop()
