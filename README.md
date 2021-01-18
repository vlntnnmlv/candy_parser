# Сandy parser :candy:

Simple product's price scrapper for the following websites:
  * https://vtk-moscow.ru/
  * https://tortomaster.ru/
  * https://bakerstore.ru/

Input data should be in the following format, contained in ".xlsx" file:

|      | bakerstore |  VTK  | tortomaster |
|------|------------|-------|-------------|
| name |    link    | link  |    link     |
| name |    link    | link  |    link     |
| name |    link    | link  |    link     |

Output ".xlsx" file bw created in near the executable, with data fromat like:

|      | bakerstore |  VTK  | tortomaster |
|------|------------|-------|-------------|
| name |    price   | price |    price    |
| name |    price   | price |    price    |
| name |    price   | price |    price    |
