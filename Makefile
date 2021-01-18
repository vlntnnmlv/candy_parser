NAME = candy_parser

$(NAME):
	pyinstaller --onefile main.py -n $(NAME) --hidden-import cmath

all: $(NAME)
	make clean
	
clean:
	rm candy_parser.spec
	rm -rf build
	rm -rf __pycache__