#! /usr/bin/env python3

import warnings
import os
import sys
import tkinter.ttk as ttk
import threading
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from bs4 import BeautifulSoup as bs
from itertools import islice
from tkinter import *
from tkinter import filedialog
from CandyExcel import CandyExcel


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
		if "permission denied" in str(e).lower():
			lbl.config(text = "Что-то пошло не так!\nПожалуйста, закройте Excel файл!\n" + str(e).replace("[Errno 13] Permission denied: ",""), fg = "red")
		else:
			lbl.config(text = "Что-то пошло не так!\nОбратитесь к разработчику!\n" + str(e), fg = "red")


def do_job(filename):
	''' Parser function. Sends requests by link in the file,
		and changes data with price values '''

	GUIParams = {"lbl" : lbl, "pb" : pb}
	pb['value'] = 0

	ce = CandyExcel()
	ce.set_raw_data(filename, GUIParams)
	ce.create_new_data(GUIParams)
	ce.save_data(GUIParams)
	ce.prettify_data(filename, GUIParams)
	lbl.config(text ="Готово!", fg="green", font=("bold"))


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
	pb = ttk.Progressbar(back, mode="determinate", maximum = 90)
	pb.pack(pady=20)
	btn = Button(back, text="Выбрать файл", command=lambda : threading.Thread(target=onclick).start())
	btn.pack()
	window.bind('<Escape>', sys.exit)
	window.mainloop()