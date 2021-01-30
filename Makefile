NAME = "Candy Parser"
ICON = candy.ico

$(NAME):
	pyinstaller --onefile -n $(NAME) --hidden-import cmath --windowed --noconsole --icon=$(ICON) main.py CandyExcel.py

all: $(NAME)
	make clean
	
clean:
	rm candy_parser.spec
	rm -rf build
	rm -rf __pycache__