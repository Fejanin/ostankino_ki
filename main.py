import moduls.worker_xlsx as W


#file = 'Крым ЛП 06,07,23кгопт-4я машина заказ№4.xlsx'
file = input('Введите название файла: ')

# чистый бланк
new_file = 'Бланк КИ  клиент ООО ЛОГИСТИЧЕСКИЙ ПАРТНЕР на отгрузку с 10.07.2023.xlsx'
    
#W.POKOM_Rewriter(file, new_file)
W.POKOM_Rewriter(file, new_file, '1С')
