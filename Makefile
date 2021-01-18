NAME = candy_parser

all:
	pyinstaller --onefile main.py -n $(NAME) --hidden-import cmath