import xlwt
"""Записывает данные в xls файл. Если файла нет - создает его.
"""

# Инизиализируем workbook
book = xlwt.Workbook(encoding="utf-8")

# Добавляем лист workbook
sheet1 = book.add_sheet("Лист 1")

# Записываем данные в ячейку
sheet1.write(0, 0, "ячейка А1")

# Сохраняем workbook
book.save("book.xls")

# доп инфа:
# https://habr.com/ru/companies/otus/articles/331998/