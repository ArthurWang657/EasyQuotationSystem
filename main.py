import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog

customerData = pd.read_excel("客戶資料.xlsx")
window = tk.Tk()
window.title("客戶系統")
window.minsize(width=1000, height=1000)
window.resizable(width=False, height=False)

def combobox_selected(event):
     print(customerCombo.current(), customerCombo.get())
     labelText.set('客戶 : ' + customerCombo.get())
     fliter = (customerData["客戶名稱"] == customerCombo.get())
     shoppingList.delete(*shoppingList.get_children())
     currentData = customerData[fliter]
     for i, row in currentData.iterrows():
         shoppingList.insert('', tk.END, values=list(row))
     print(currentData)

def treeViewClick(event):
    item = shoppingList.selection()[0]
    itemValue = shoppingList.item(item, "values")
    shoppingList.item(item, values=(itemValue[0], itemValue[1], itemValue[2], itemValue[3], '0'))
    print("Text : ", item, "    Value : ", itemValue[4])
    col_id = itemValue[1]

    # 創建一個對話框，允許用戶輸入新數據
    new_data = simpledialog.askstring("編輯數據", f"{col_id} 出貨數:", parent=window)

    if new_data is not None:
        # 更新TreeView中的數據
        shoppingList.item(item, values=(itemValue[0], itemValue[1], itemValue[2], itemValue[3], new_data))

def showFinalResult():
    # 清空TreeView
    quotation.delete(*quotation.get_children())
    for item in shoppingList.get_children():
        itemValue = shoppingList.item(item, "values")
        if itemValue[4] is not None and itemValue[4] != "nan" :
            calculated_amount = float(itemValue[2]) * float(itemValue[3]) * float(itemValue[4])  # 假設金額是出貨數*2
            itemValue += (calculated_amount,)
            quotation.insert('', 'end', values=itemValue)
            print("Final Value : ", itemValue)

customerNames = pd.read_excel("客戶資料.xlsx", usecols=['客戶名稱'])
column_names = ['客戶名稱']
customerList = []
customerNames = customerNames.drop_duplicates()
customerNames = customerNames.reset_index(drop=True)
customerList = list(customerNames['客戶名稱'].tolist())

labelTop = tk.Label(window, text="選擇客戶", height=2, font=('微軟正黑體', 12))
labelTop.pack()

comboboxText = tk.StringVar()
customerCombo = ttk.Combobox(window, values=customerList, height=10, state='readonly', font=('微軟正黑體', 12))
customerCombo.current(0)
customerCombo.pack()
customerCombo.bind('<<ComboboxSelected>>', combobox_selected)

labelText = tk.StringVar()
customerLabel = tk.Label(window, textvariable=labelText, height=5, font=('Arial', 12))
customerLabel.pack()

shoppingList = ttk.Treeview(window, columns=list(customerData.columns), show='headings')
for col in customerData.columns:
    shoppingList.heading(col, text=col)

shoppingList.bind('<ButtonRelease-1>', treeViewClick)
shoppingList.pack()

quotationList = []
show_button = tk.Button(window, text="結算", height=2, command=showFinalResult)
listTitle = list(customerData.columns)
listTitle.append("金額")
quotation = ttk.Treeview(window, columns=listTitle , show='headings')
show_button.pack()
quotation.pack()

window.mainloop()
