from openpyxl import load_workbook
""" Работа с exel файлами
С фриланса:
Нужно написать скрипт, который ищет пустые ячейки и заменяет их на определенные значения

"""

wb = load_workbook('sber_tmp.xlsx')
ws = wb['Лист1']

def search_empty_cells():
    """ функция ищет пустые значения и применяет к ним функцию insert_in_cells"""
    for i in ws:
        for j in i:
            cell = str(j)[-3:-1] # имя ячейки в str формате
            print('cell = ',cell, ws[cell].value)
            if ws[cell].value == None:
                if cell[0] == 'B':
                    insert_into_cell(cell, 'пустое имя')
                elif cell[0] == 'C':
                    insert_into_cell(cell, 'нет фамилии')
                elif cell[0] == 'D':
                    insert_into_cell(cell, 'нет телефона')
                else:
                    insert_into_cell(cell, 'нет почты')

# (<Cell 'Лист1'.A1>, <Cell 'Лист1'.B1>, <Cell 'Лист1'.C1>, <Cell 'Лист1'.D1>, <Cell 'Лист1'.E1>)

def find_index(cell1, cell2):
    second_word = cell1.split(' ')[1]
    ind = 0
    for i in cell2.split(' '):
        if i == second_word:
            return ind
        ind += 1



def compare(num: int) -> bool:
    num_cell_B = 'B' + str(num+1)
    num_cell_C = 'C' + str(num+1)
    tmp1 = ws[num_cell_B].value
    tmp2 = ws[num_cell_C].value
    swi =  find_index(tmp1, tmp2)
    print(tmp1.split(' ')[1:])
    print(tmp2.split(' ')[swi:])
    print()



def insert_into_cell(cell: str, data: str):
    """ вставляет в ячейку cell значение data"""
    ws[cell] = data
    for i in ws.values:
        print(i)
    wb.save('exel_table1.xlsx')

if __name__ == '__main__':
    compare(284)


    # search_empty_cells()
