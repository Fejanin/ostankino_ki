import moduls.worker_xlsx as W


#file = 'Заказ на 10.07 2.xlsx'
#file = 'чистый бланк ПОКОМ.xlsx'
file = 'Крым ЛП 06,07,23кгопт-4я машина заказ№4.xlsx'
new_file = 'чистый бланк ПОКОМ.xlsx'
'''
obj1 = W.POKOM_Reader(file)
for i in obj1.all_rows:
    print(i)
print(obj1())
'''
#obj2 = W.POKOM_Rewriter(file, new_file)
#print(obj2.write_file_name)
#print(obj2.read_file())
    
obj = W.POKOM_Rewriter(file, new_file)
