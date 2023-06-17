import locale
import tkinter
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Treeview
from datetime import date, datetime
import sqlite3
from docxtpl import DocxTemplate
from num2words import num2words
import calendar


# creating table
# cursor.execute("CREATE TABLE addresses (name text, code text, pvm_code text, address text)")
def update_list():
    conn = sqlite3.connect('client_list.db')
    cursor = conn.cursor()
    cursor.execute("SELECT *, oid FROM addresses")
    global client_list
    client_list = cursor.fetchall()
    conn.commit()
    conn.close()

update_list()

today = date.today()
no = 0





window = Tk()
window.title("Sąskaitų forma")
frame = Frame(window)
frame.pack(padx=20, pady=10)

def new_window():

    newWindow = Toplevel(window)
    newWindow.title("Klientų sąrašas")
    frame2 = Frame(newWindow)
    frame2.pack(padx=30, pady=10)

    # def save_clients():
    #     print("000000000", client_list)
    #     with open('Klientai.txt', 'w') as filehandle:
    #         for listitem in client_list:
    #             print("preš rašymą", listitem, client_list)
    #             filehandle.write(str(listitem)+"\n")
    #     newWindow.destroy()
    # newWindow.protocol('WM_DELETE_WINDOW', save_clients)

    def delete_selected():
        selected_item = tree2.selection()[0]
        item = tree2.item(selected_item)
        values = item.get("values")
        print(values)
        print(values[1])
        tree2.delete(selected_item)


        conn = sqlite3.connect('client_list.db')
        cursor = conn.cursor()
        cursor.execute("DELETE from addresses WHERE code='%s'" % values[1])

        conn.commit()
        conn.close()
        update_list()
        client_drop['menu'].delete(0, 'end')
        for i in client_list:
            client_drop['menu'].add_command(label=i, command=tkinter._setit(client, i))
        client.set(client_list[0])


    def clear_client():
        description.delete(0, tkinter.END)
        code_entry.delete(0, tkinter.END)
        pvm_entry.delete(0, tkinter.END)
        address_entry.delete(0, tkinter.END)

    def add_client():
        description2 = description.get()
        code_entry2 = code_entry.get()
        pvm_entry2 = pvm_entry.get()
        address_entry2 = address_entry.get()
        client_item = [description2, code_entry2, pvm_entry2, address_entry2]
        # client_list.append(client_item)
        # Create a database
        conn = sqlite3.connect('client_list.db')
        cursor = conn.cursor()
        cursor.execute("INSERT INTO addresses VALUES(:description2, :code_entry2, :pvm_entry2, :address_entry2)",
                       {
                        'description2': description2,
                        'code_entry2': code_entry2,
                        'pvm_entry2': pvm_entry2,
                        'address_entry2': address_entry2
                       })


        conn.commit()
        conn.close()
        update_list()

        tree2.insert('', tkinter.END, values=client_item)
        clear_client()
        client_drop['menu'].delete(0, 'end')
        for i in client_list:
            client_drop['menu'].add_command(label=i, command=tkinter._setit(client, i))
        client.set(client_list[0])



    columns2 = ('name', 'code', 'pvm', 'address')
    tree2 = Treeview(frame2, columns=columns2, show="headings")
    tree2.column('name', width=130)
    tree2.column('code', width=110)
    tree2.column('pvm', width=110)
    tree2.column('address', width=180)
    tree2.heading('name', text='Pavadinimas')
    tree2.heading('code', text='įmonės kodas')
    tree2.heading('pvm', text='Pvm kodas')
    tree2.heading('address', text="adresas")
    tree2.grid(row=0, rowspan=12, column=0, columnspan=5, padx=10, pady=10)
    for i in client_list:
        tree2.insert("", "end", values=i)

    description_label = Label(frame2, text="Pavadinimas")
    description_label.grid(row=2, column=6)
    description = Entry(frame2, width=32)
    description.grid(row=3, column=6)

    code_label = Label(frame2, text="Įmonės kodas")
    code_label.grid(row=4, column=6)
    code_entry = Entry(frame2, width=32)
    code_entry.grid(row=5, column=6)

    pvm_label = Label(frame2, text="Pvm mokėtojo kodas")
    pvm_label.grid(row=6, column=6)
    pvm_entry = Entry(frame2, width=32)
    pvm_entry.grid(row=7, column=6)

    address_label = Label(frame2, text="Adresas")
    address_label.grid(row=8, column=6)
    address_entry = Entry(frame2, width=32)
    address_entry.grid(row=9, column=6)

    add_button = Button(frame2, text="Pridėti", command= add_client)
    add_button.grid(row=10, column=6, pady=10, sticky="ew")

    del_button = Button(frame2, text="Ištrinti", command=delete_selected)
    del_button.grid(row=11, column=6, pady=15, sticky="ew")



def clear_item():
    description.delete(0, tkinter.END)
    qty.delete(0, tkinter.END)
    qty.insert(0, "1")
    price.delete(0, tkinter.END)
    price.insert(0, "0.0")


def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    invoice_year2 = date_entry.get()
    invoice_nr2 = invoice_nr.get()
    car2 = car_entry.get()
    sum2 = sum(float(item[5]) for item in invoice_list)
    pvm2 = sum2 * 0.21
    client_list = client.get().replace("\'", "").strip('()').split(', ')
    total = sum2 + pvm2
    numbers2words = num2words(total, to = 'currency', lang='lt')
    date = datetime.strptime(invoice_year2, '%Y-%m-%d')
    print(date)
    locale.setlocale(locale.LC_ALL, 'LT')
    month = date.strftime("%B")
    date2 = str(date.year) + " m. " + str(month) + " m. " + str(date.day) + " d."

    doc.render({"invoice_year": invoice_year2[2:4],
                "invoice_nr": invoice_nr2,
                "date": date2,
                "car": car2,
                "invoice_list": invoice_list,
                "sum": "{:.2f}".format(sum2),
                "pvm": "{:.2f}".format(pvm2),
                "total": "{:.2f}".format(total),
                "company_name": client_list[0],
                "company_code": client_list[1],
                "pvm_code": client_list[2],
                "address": client_list[3],
                "sum_in_words": numbers2words
                })
    doc.save("Sąskaita " + client_list[0]+ " " + invoice_year2 + ".docx")
    messagebox.showinfo("Informacija", "Sąskaita "+client_list[0]+ " " + invoice_year2 + ".docx " +"sukurta")

invoice_list = []
def add_item():
    global no
    description2 = description.get()
    qty2 = int(qty.get())
    price2 = float(price.get())
    sum = "{:.2f}".format(qty2 * price2)
    no += 1
    invoice_item = [no, description2, "vnt", qty2, price2, sum]
    tree.insert('', tkinter.END, values=invoice_item)
    clear_item()
    invoice_list.append(invoice_item)

def new_invoice():
    car_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()




description_label = Label(frame, text="Pavadinimas")
description_label.grid(row=2, column=0)
description = Entry(frame, width=27)
description.grid(row=3,column=0)

qty_label = Label(frame, text="Kiekis")
qty_label.grid(row=4, column=0)
qty = Spinbox(frame, from_=1, to=2000, increment=1)
qty.grid(row=5,column=0, padx=30)

price_label = Label(frame, text="Kaina €")
price_label.grid(row=2, column=1)
price = Spinbox(frame, from_=0.0, to=15000, increment=1, format='%1.2f')
price.grid(row=3,column=1)

add_button = Button(frame, text="Pridėti prekę", command=add_item)
add_button.grid(row=5,column=1)

date_label = Label(frame, text="Data")
date_label.grid(row=0, column=2)
date_entry = Entry(frame)
date_entry.insert(0, today)
date_entry.grid(row=1,column=2)

invoice_nr_label = Label(frame, text="Sąskaitos nr.")
invoice_nr_label.grid(row=0, column=3)
invoice_nr = Entry(frame)
invoice_nr.grid(row=1,column=3)

client_label = Label(frame, text="Klientas")
client_label.grid(row=2, column=2, columnspan=3)
client = StringVar()


if client_list:
    client.set(client_list[0])
else:
    client.set(" ")
    client_list.append("['Test']")

client_drop = OptionMenu(frame, client, *client_list)
client_drop.grid(row=3, column=2, columnspan=3, padx=30)
client_drop.config(width=50)

car_label = Label(frame, text="Mašinos modelis ir numeris")
car_label.grid(row=4, column=2)
car_entry = Entry(frame, width=27)
car_entry.grid(row=5, column=2)

add_client = Button(frame, text="Pridėti naują klientą", command=new_window)
add_client.grid(row=5, column=3)


columns = ('eil', 'pav', 'vnt', 'kiek', 'kain', 'sum')
tree = Treeview(frame, columns=columns, show="headings")
tree.column("eil", width=60)
tree.column("vnt", width=60)
tree.column("pav", width=300)
tree.column("kiek", width=100)
tree.column("kain", width=100)
tree.column("sum", width=100)
tree.heading('eil', text='Eilė')
tree.heading('pav', text='Pavadinimas')
tree.heading('vnt', text='Vnt')
tree.heading('kiek', text="Kiekis")
tree.heading('kain', text="Kaina")
tree.heading('sum', text="Suma")
tree.grid(row=6, column=0, columnspan=4, padx=20, pady=10)


save_invoice_button = Button(frame, text="Generuoti sąskaitą", command=generate_invoice)
save_invoice_button.grid(row=7, column=0, columnspan=4, sticky="news", padx=20, pady=5)
new_invoice_button = Button(frame, text="išvalyti", command=new_invoice)
new_invoice_button.grid(row=8, column=0, columnspan=4, sticky="news", padx=20, pady=5)


# doc = DocxTemplate("invoice_template.docx")
# doc.render({"sum":"23424","invoice_list":invoice_list})
# doc.save("new_invoice.docx")




window.mainloop()