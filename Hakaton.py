from tkinter import *
from xlwings import *
def BT1():
    li = listbox3.curselection()
    k = li[0]
    skidka = bk.range(k+2, 2).value
    preis = bk.range(k+2, 3).value
    if text1.get(1.0):
        num = int(text1.get(1.0))
    else:
        num = 0
    if skidka == 'нет':
        skidka = 0
    total_preis = preis*(100-skidka)/100
    label_1 = Label(root, width=30, font='14', text=f'Наименование товара: {bk.range(k+2, 1).value}', wraplength=170)
    label_2 = Label(root, width=30, font='14', text=f'Цена товара: {preis}')
    label_3 = Label(root, width=30, font='14', text=f'Скидка: {skidka}%')
    label_4 = Label(root, width=30, font='14', text=f'Цена со скидкой: {total_preis}')
    label_5 = Label(root, width=30, font='14', text=f'Количество товаров: {num}')
    label_6 = Label(root, width=30, font='14', text=f'Общая стоимость: {total_preis*num//1}')
    label_1.grid(row=2, column=1)
    label_2.grid(row=3, column=1)
    label_3.grid(row=4, column=1)
    label_4.grid(row=5, column=1)
    label_5.grid(row=6, column=1)
    label_6.grid(row=7, column=1)
    Button1['state'] = 'disabled'
bk = Book('Товары.xlsx').sheets("Движение товаров")
root = Tk()
root.geometry('600x800')
listbox3=Listbox(root,height=10, width=30, selectmode=SINGLE, font='14')
for i in bk.range((2,1), (99999, 1)).value:
    if i == None:
        break
    listbox3.insert(END, i)
Label1 = Label(root, text='Введите количество товаров', font='14', width=30, height=10)
Button1= Button(root, text='Посчитать', command=BT1, height=10)
text1 = Text(root, width=30, font='14',height=10)
Label1.grid(row = 1, column=2)
text1.grid(row = 2, column=2, rowspan=6, sticky='ns')
listbox3.grid(row=1, column=1)
Button1.grid(row=8, column=1, columnspan=2, sticky='we')
root.mainloop()

