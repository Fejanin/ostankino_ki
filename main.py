import moduls.worker_xlsx as W


#file = 'Крым ЛП 06,07,23кгопт-4я машина заказ№4.xlsx'
file = input('Введите название файла: ')

# чистый бланк
new_file = 'Бланк КИ  клиент ООО ЛОГИСТИЧЕСКИЙ ПАРТНЕР на отгрузку с 10.07.2023.xlsx'

while True:
    type_file = input('Введите 1, если данные переносятся из старогог файла ПОКОМа или 2, если файл другого образца: ')
    if type_file != '1' and type_file != '2':
        print('Принимается только значение 1 или 2!')
    else:
        break
if type_file == '1':
    W.POKOM_Rewriter(file, new_file)
else:
    W.POKOM_Rewriter(file, new_file, '1С')
