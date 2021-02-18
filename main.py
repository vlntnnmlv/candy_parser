#! /usr/bin/env python3

import warnings
import os
import sys
import tkinter.ttk as ttk
import threading
import openpyxl
from openpyxl.descriptors import base
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from bs4 import BeautifulSoup as bs
from itertools import islice
from tkinter import *
from tkinter import filedialog
from candyexcel import CandyExcel
import candyexcel
import time
from ntpath import basename

G = ""

def onclick(event = None):
	''' Checks is chosen file valid and calls parser function '''
	try:
		res = filedialog.askopenfilename(
			initialdir = "/",
			title = "Select file",
			filetypes = (("Excel files","*.xlsx"),("all files","*.*"))
			)
		if (res):
			exec_path = res
			path.config(text = "PATH: " + os.path.abspath(res))
			if (res.split(".")[-1] != "xlsx"):
				lbl.config(text = "Введите название Ексель файла.", fg="red")
			elif not os.path.exists(res):
				lbl.config(text ="Такого Ексель файла не сущесвует!", fg="red")
			else:
				btn.config(state=DISABLED)
				do_job(res)
				btn.config(state=NORMAL)
	except BaseException as e:
		if "permission denied" in str(e).lower():
			lbl.config(
				text =
					"Что-то пошло не так!\
					\nПожалуйста, закройте Excel файл!\
					\n" + str(e).replace("[Errno 13] Permission denied: ",""),
				fg = "red"
				)
		else:
			lbl.config(
				text =
					"Что-то пошло не так!\
					\nОбратитесь к разработчику!\
					\n" + str(e),
				fg = "red"
				)

def onclick2(event = None):
	''' Checks is chosen file valid and calls mailing function '''
	try:
		if getattr(sys, 'frozen', False):
			application_path = os.path.dirname(sys.executable)
		elif __file__:
			application_path = os.path.dirname(__file__)
		res = os.path.join(application_path, "emails.xlsx")
		if (res):
			exec_path = res
			path.config(text = "PATH: " + os.path.abspath(res))
			if (res.split(".")[-1] != "xlsx"):
				lbl.config(text = "Введите название Ексель файла.", fg="red")
			elif not os.path.exists(res):
				lbl.config(text ="Ексель файла для рассылки не сущесвует!", fg="red")
			else:
				btn.config(state=DISABLED)
				do_job2(res)
				btn.config(state=NORMAL)
	except BaseException as e:
		if "permission denied" in str(e).lower():
			lbl.config(
				text =
					"Что-то пошло не так!\
					\nПожалуйста, закройте Excel файл!\
					\n" + str(e).replace("[Errno 13] Permission denied: ",""),
				fg = "red"
				)
		else:
			lbl.config(
				text =
					"Что-то пошло не так!\
					\nОбратитесь к разработчику!\
					\n" + str(e),
				fg = "red"
				)

def do_job(filename):
	''' Parser function. Sends requests by link in the file,
		and changes data with price values '''
	pb['value'] = 0
	ce = CandyExcel()
	pb.start(10)
	lbl.config(text = "Загружаем данные...")
	ce.clone_update(filename)

	if (os.path.exists(basename(filename).split('.')[0] + "_out.xlsx")):
		lbl.config(text = "Рассчитываем изменения...")
		ce.calc_changes(filename)

	ce.prettify()
	ce.close_data(filename)
	pb.stop()
	lbl.config(text ="Готово!", fg="green", font=("bold"))
	send.config(state = NORMAL)

	global G
	G = basename(filename).split('.')[0] + "_out.xlsx"

def do_job2(filename):
	global G
	send.config(state = DISABLED)
	lbl.config(text = "Проводим рассылку...", fg="blue")
	pb.start(10)
	# try:
	print("!1\n")
	candyexcel.mailing(filename, G)
	print("!2\n")
	# except:
	# 	lbl.text = "Что-то пошло не так! Попробуйте еще раз!"
	# 	pb.stop()
	# 	return
	pb.stop()
	lbl.config(text ="Готово!", fg="green", font=("bold"))
	send.config(state = NORMAL)

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
	
	pb = ttk.Progressbar(back, mode="indeterminate")
	pb.value = 0
	pb.pack(pady=20)
	
	btn = Button(back, text="Выбрать файл", command=lambda : threading.Thread(target=onclick).start())
	btn.pack()

	send = Button(back, text="Разослать", command=lambda : threading.Thread(target=onclick2).start())
	send.pack()
	send.config(state = DISABLED)
	
	window.bind('<Escape>', sys.exit)
	window.mainloop()

