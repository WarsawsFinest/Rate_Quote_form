import csv
import datetime as dt
import sqlite3
from tkinter import *
from tkinter import ttk

root = Tk()
root.geometry("700x700")
root.title("TMC Quote Form")
# root.configure(bg='#8D734A')
header = Label(text="Quote Form", bg="#F26122", font=('ITC Franklin Gothic Std', 12), fg="black", width='500', height='3')
header.pack()

# Create a db or connect to one
conn = sqlite3.connect('quotedrates.db')

# Create cursor
c = conn.cursor()
# Drop Table
# c.execute('DROP TABLE quoted_rates')

# Create Table

c.execute("""CREATE TABLE IF NOT EXISTS quotedrates (date integer,
        lsr_name text,
        office text,
        gm text,
        bill_to_code integer,
        origin_zip integer,
        dest_zip integer,
        comments text,
        truck text,
        quoted_rate real,
        linehaul real,
        fsc real,
        accessorials real,
        eqt_type text,
        spec_req_combo text,
        spec_req_combo2 text,
        spec_req_combo3 text,
        od integer,
        weight integer,
        rate_source text
        )""")

# Create Date Entry Widget

d = Entry(root)
d.place(x=105, y=70)
d.insert(0, f"{dt.datetime.now():%a %m/%d/%Y}")


def dateselector():
    date = Label(root, text=d.get())
    date.place(x=90, y=70)


# Create CSV Headers
def create_csv_headers():
    with open('quoterates.csv', 'w', newline='') as f:
        h = csv.DictWriter(f, fieldnames=['Date', 'LSR Name', 'Office', 'GM', 'Bill to Code', 'Origin Zip',
                                          'Destination Zip', 'Comments', 'Quoted Rate', 'Linehaul', 'FSC',
                                          'Accessorial', 'OD', 'Weight', 'Truck Type', 'Equipment Type',
                                          'Special Requirements', 'Special Req1', 'Special Req2', 'Rate Source', 'Quote ID'])
        h.writeheader()


create_csv_headers()


# write to csv excel
def write_to_csv(result):
    with open("quoterates.csv", 'a', newline='') as f:
        w = csv.writer(f, dialect="excel")
        for record in result:
            w.writerow(record)


# save and export to excel
def save_quotes():
    save_quotes_query = Tk()
    save_quotes_query.title("All Quotes")
    save_quotes_query.geometry("100x100")
    # Query db
    # Create a db or connect to one
    conn = sqlite3.connect('quotedrates.db')

    # Create cursor
    c = conn.cursor()
    c.execute("SELECT *, oid FROM quotedrates")
    result = c.fetchall()
    for index, x in enumerate(result):
        num = 0
        for y in x:
            num += 1
    # # Commit Changes
    conn.commit()

    # Close connection
    conn.close()
    csv_button = Button(save_quotes_query, text="Save to Excel", command=lambda: write_to_csv(result))
    csv_button.pack()




# # Create a function to delete a record
# def delete():
#     # Create a db or connect to one
#     conn = sqlite3.connect('quotedrates.db')
#
#     # Create cursor
#     c = conn.cursor()
#
#     # Delete a record
#     c.execute("DELETE from quotedrates WHERE oid = " + delete_box.get())
#
#     # Commit Changes
#     conn.commit()
#
#     # Close connection
#     conn.close()


# Create Submit Function For Database

def submit():
    # Create a db or connect to one
    conn = sqlite3.connect('quotedrates.db')

    # Create cursor
    c = conn.cursor()

    # Insert Into Table
    c.execute(
        "INSERT INTO quotedrates VALUES (:date, :lsr_name, :office, :gm, :bill_to_code, :origin_zip, :dest_zip, "
        ":quoted_rate, :linehaul, :fsc, :accessorial, :od, :weight, :truck_select, :eqt_select, "
        ":spec_req_combo,:spec_req_combo2, :spec_req_combo3, :source_select, :comments,)",
        {
            'date': d.get(),
            'lsr_name': lsr_name.get(),
            'office': office.get(),
            'gm': gm.get(),
            'bill_to_code': bill_to_code.get(),
            'origin_zip': origin_zip.get(),
            'dest_zip': dest_zip.get(),
            'quoted_rate': quoted_rate.get(),
            'linehaul': linehaul.get(),
            'fsc': fsc.get(),
            'accessorial': accessorial.get(),
            'od': od.get(),
            'weight': weight.get(),
            'truck_select': truck_select.get(),
            'eqt_select': eqt_select.get(),
            'spec_req_combo': spec_req_combo.get(),
            'spec_req_combo2': spec_req_combo2.get(),
            'spec_req_combo3': spec_req_combo3.get(),
            'source_select': source_select.get(),
            'comments': comments.get()
        })

    # Clear Text boxes
    lsr_name.delete(0, END)
    office.delete(0, END)
    gm.delete(0, END)
    bill_to_code.delete(0, END)
    origin_zip.delete(0, END)
    dest_zip.delete(0, END)
    comments.delete(0, END)
    quoted_rate.delete(0, END)
    linehaul.delete(0, END)
    fsc.delete(0, END)
    accessorial.delete(0, END)
    od.delete(0, END)
    weight.delete(0, END)
    truck_select.delete(0, END)
    eqt_select.delete(0, END)
    spec_req_combo.delete(0, END)
    spec_req_combo2.delete(0, END)
    spec_req_combo3.delete(0, END)
    source_select.delete(0, END)

    # # Commit Changes
    conn.commit()

    # Close connection
    conn.close()


# Create text boxes
date = Label(root)
date.place(x=100, y=70)
lsr_name = Entry(root)
lsr_name.place(x=100, y=110)
office = Entry(root)
office.place(x=100, y=150)
gm = Entry(root)
gm.place(x=100, y=190)
bill_to_code = Entry(root)
bill_to_code.place(x=100, y=230)
origin_zip = Entry(root)
origin_zip.place(x=100, y=270)
dest_zip = Entry(root)
dest_zip.place(x=100, y=310)
comments = Entry(root)
comments.place(x=100, y=510)
quoted_rate = Entry(root)
quoted_rate.place(x=100, y=350)
linehaul = Entry(root)
linehaul.place(x=100, y=390)
fsc = Entry(root)
fsc.place(x=100, y=430)
accessorial = Entry(root)
accessorial.place(x=100, y=470)
od = Entry(root)
od.place(x=500, y=375)
weight = Entry(root)
weight.place(x=500, y=410)
# delete_box = Entry(root, width=10)
# delete_box.place(x=500, y=600)

# Create text box labels
date_label = Label(root, text="Date").place(x=20, y=70)
lsr_name_label = Label(root, text="LSR Name ", ).place(x=20, y=110)
office_label = Label(root, text="TMC Office * ", ).place(x=20, y=150)
gm_label = Label(root, text="GM * ", ).place(x=20, y=190)
bill_to_code_label = Label(root, text="Bill to Code *", ).place(x=20, y=230)
origin_zip_label = Label(root, text="Origin Zip *", ).place(x=20, y=270)
dest_zip_label = Label(root, text="Dest Zip *", ).place(x=20, y=310)
quoted_rate_label = Label(root, text="Quoted Rate *", ).place(x=20, y=350)
linehaul_label = Label(root, text="Linehaul Rate").place(x=20, y=390)
fsc_label = Label(root, text="FSC").place(x=20, y=430)
accessorial_label = Label(root, text="Accessorials").place(x=20, y=470)
comments_label = Label(root, text="Comments *", ).place(x=20, y=510)
od_label = Label(root, text="OD: LxWxH").place(x=400, y=375)
weight_label = Label(root, text="Weight").place(x=400, y=410)
truck_label = Label(root, text="Truck Type").place(x=420, y=100)
equipment_label = Label(root, text="Equipment Type").place(x=395, y=150)
special_requirements_label = Label(root, text="Special Requirements").place(x=370, y=250)
rate_source_label = Label(root, text="Rate Source").place(x=400, y=200)

# delete_box_label = Label(root, text="Select ID")
# delete_box_label.place(x=500, y=575)


truck = StringVar()
truck_select = ttk.Combobox(root, width=25, textvariable=truck)
truck_select['values'] = ('', 'LTL', 'PTL', 'FTL')
truck_select.place(x=490, y=100)
truck_select.current(0)

eqt_type = StringVar()
eqt_select = ttk.Combobox(root, width=25, textvariable=eqt_type)
eqt_select['values'] = ('',
                        'Reefer',
                        'flatbed',
                        'RGN',
                        'Van',
                        'Drayage',
                        'Intermodal',
                        'Stepdeck')
eqt_select.place(x=490, y=150)
eqt_select.current(0)

rate_source = StringVar()
source_select = ttk.Combobox(root, width=25, textvariable=rate_source)
source_select['values'] = ('', 'Carrier', 'DAT', 'ITS', 'Co-Worker', 'Other LSR', 'Other CSS')
source_select.place(x=490, y=200)
source_select.current(0)


def comboclick(event):
    spec_req_label = Label(root, text=spec_req_combo.get())
    spec_req_label.pack(padx=490, pady=250)
    spec_req_label = Label(root, text=spec_req_combo2.get())
    spec_req_label.pack(padx=490, pady=275)
    spec_req_label = Label(root, text=spec_req_combo3.get())
    spec_req_label.pack(padx=490, pady=300)


special_requirements = ['', 'HC', 'Tanker', 'OD', 'Tarp', 'Temp Control', 'Pallet Jack', 'Lumper',
                        'Crane and Rigging',
                        'Weekend', 'Pre-Pay', 'Multistop', 'US Citizen', 'Pallet Exchange', 'Drop/Hook',
                        'Overweight', 'Bobtail',
                        'Scale Light/Heavy', 'Driver Assist', 'Hazmat', 'Re-Tarp', 'Jobsite Delivery',
                        'Blind', 'Expedited Load Surcharge',
                        'Burroughs Charge']
spec_req_combo = ttk.Combobox(root, width=25, value=special_requirements)
spec_req_combo.current(0)
spec_req_combo.bind("<<ComboboxSelected>>", comboclick)
spec_req_combo.place(x=490, y=250)
spec_req_combo2 = ttk.Combobox(root, width=25, value=special_requirements)
spec_req_combo2.current(0)
spec_req_combo2.bind("<<ComboboxSelected>>", comboclick)
spec_req_combo2.place(x=490, y=275)
spec_req_combo3 = ttk.Combobox(root, width=25, value=special_requirements)
spec_req_combo3.current(0)
spec_req_combo3.bind("<<ComboboxSelected>>", comboclick)
spec_req_combo3.place(x=490, y=300)

# Create a submit button
submit_btn = Button(root, text="Submit (Click First)", width="100", font=('ITC Franklin Gothic Std',9), height="2", command=submit, bg="#231F20", fg="white")
submit_btn.place(x=0, y=610)

# Create a delete Button
# delete_btn = Button(root, text="Delete Record", command=delete)
# delete_btn.place(x=400, y=600)

# button to bring up excel save button
list_quotes = Button(root, text='Click to open "Save to Excel" Button',font=('ITC Franklin Gothic Std',9),  width="100", height="2", command=save_quotes, bg="#F26122", fg="black", )
list_quotes.place(x=0, y=650)

# Commit Changes
conn.commit()

# Close connection
conn.close()

root.mainloop()
