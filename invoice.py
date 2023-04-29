import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
import numpy as np
import random

window=tkinter.Tk()
window.config(bg="#8B8B86")

invoice_list=[]
def add_item():
    qty=int(qnt_spinbox.get())
    desc=desc_entry.get()
    price=float(price_spin.get())
    line_total=qty*price
    invoice_item=[desc,qty,price,line_total]
    tree.insert('',0,values=invoice_item)
    clear_item()

    invoice_list.append(invoice_item)

def clear_item():
    qnt_spinbox.delete(0,tkinter.END)
    qnt_spinbox.insert(0,"1")
    desc_entry.delete(0,tkinter.END)
    price_spin.delete(0,tkinter.END)
    price_spin.insert(0,"0.0")

def new_invoice():
    first_name_entry.delete(0,tkinter.END)
    last_name_entry.delete(0,tkinter.END)
    phone_entry.delete(0,tkinter.END)
    tree.delete(*tree.get_children())

    invoice_list.clear()

def generate_invoice():
    doc=DocxTemplate("C:\\Users\\ARWINDD\\ETP\\Project\\invoice_template.docx")
    name=first_name_entry.get()+" "+last_name_entry.get()
    phone=phone_entry.get()
    subtotal=sum(item[3] for item in invoice_list)
    discount=0.06
    gst=0.18
    invoice_number=random.randint(1000,100000)
    date=datetime.datetime.now().strftime("%Y-%m-%d")
    total=subtotal*(1+gst)-subtotal*discount

    doc.render({"name":name,
               "phone":phone,
                "invoice":invoice_number,
                "invoice_list":invoice_list,
                "date":date,
                "subtotal":subtotal,
                'discount':str(discount*100)+"%",
                'GST':str(gst*100)+"%",
                "total":round(total,2)})
    doc_name="new_invoice"+name+".docx"
    doc.save(doc_name)
    messagebox.showinfo("Invoice Complete","Invoice Generated")
    new_invoice()





window.title("Invoice_Generator_Form")
frame1=tkinter.Frame(window,bg="#4a7a8c")
frame1.pack(padx=20,pady=10)
lbl=tkinter.Label(frame1,text="Get Your Invoice",font=('Bell MT',30,'bold'),bg="#4a7a8c")
lbl.pack()

frame=tkinter.Frame(frame1)
frame.pack(padx=20,pady=10)
frame.config(bg="#116562")

first_name_lbl=tkinter.Label(frame,text="First Name",fg='black',bg="#116562",font=('Bookman Old Style',14))
first_name_lbl.grid(row=0,column=0)
first_name_entry=tkinter.Entry(frame)
first_name_entry.grid(row=1,column=0)

last_name_lbl=tkinter.Label(frame,text="Last Name",fg='black',bg="#116562",font=('Bookman Old Style',14))
last_name_lbl.grid(row=0,column=1)
last_name_entry=tkinter.Entry(frame)
last_name_entry.grid(row=1,column=1)

phone_lbl=tkinter.Label(frame,text="Phone",fg='black',bg="#116562",font=('Bookman Old Style',14))
phone_lbl.grid(row=0,column=2)
phone_entry=tkinter.Entry(frame)
phone_entry.grid(row=1,column=2)

qnt_label=tkinter.Label(frame,text="Qty",fg='black',bg="#116562",font=('Bookman Old Style',14))
qnt_label.grid(row=3,column=1)
qnt_spinbox=tkinter.Spinbox(frame,from_=1,to=100)
qnt_spinbox.grid(row=4,column=1)

Description_label=tkinter.Label(frame,text="Description",fg='black',bg="#116562",font=('Bookman Old Style',14))
Description_label.grid(row=3,column=0)
desc_entry=tkinter.Entry(frame)
desc_entry.grid(row=4,column=0)

Unit_price_label=tkinter.Label(frame,text="Unit Price",fg='black',bg="#116562",font=('Bookman Old Style',14))
Unit_price_label.grid(row=3,column=2)
price_spin=tkinter.Spinbox(frame,from_=0.0,to=500,increment=0.5)
price_spin.grid(row=4,column=2)


add_item_btn=tkinter.Button(frame,text="Add Item",command=add_item,fg='white',bg="#82C1CF",font=('Goudy Old Style',14))
add_item_btn.grid(row=5,column=2,pady=5)



columns=('desc','qty','price','total')
tree=ttk.Treeview(frame,columns=columns,show="headings")
tree.heading('qty',text="Qty")
tree.heading('desc',text="Description")
tree.heading('price',text="Unit Price")
tree.heading('total',text="Total")
tree.grid(row=6,column=0,columnspan=3,padx=20,pady=10)


save_invoice_btn=tkinter.Button(frame,text="Generate Invoice",command=generate_invoice,bg='black',fg='white',font=('Bell MT',14))
save_invoice_btn.grid(row=7,column=0,columnspan=3,sticky="news",padx=20,pady=5)
new_invoice_btn=tkinter.Button(frame,text="New Invoice",command=new_invoice,bg='white',fg='black',font=('Bell MT',14))
new_invoice_btn.grid(row=8,column=0,columnspan=3,sticky="news",padx=20,pady=5)






# window.geometry("600x440")
window.mainloop()
