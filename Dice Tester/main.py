import keyboard
import xlwt

style0 = xlwt.easyxf('font: name Arial, color-index black, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf ('font: name Arial, color-index black, bold off')

wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')
wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')

ws.write(0, 0, "#", style0)
ws.write(0, 1, "Numero 1", style0)
ws.write(0, 2, "Numero 2", style0)
ws.write(0, 3, "Numero 3", style0)

n1 = 0
n2 = 0
n3 = 0 
while True: #111223333331111222333312311111111111122222222223333333333333

    
    if keyboard.read_key() == "1":
        n1 = n1 + 1
        print (n1)
    if keyboard.read_key() == "2":
        n2 = n2 + 1
        print (n2)
    if keyboard.read_key() == "3":
        n3 = n3 + 1
        print (n3)
    if keyboard.read_key() == "enter":
        ws.write (1, 1, n1, style1)
        ws.write (1, 2, n2, style1)
        ws.write (1, 3, n3, style1)
        wb.save('example.xls')
        break