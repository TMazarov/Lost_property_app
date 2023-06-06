import tkinter as tk
from tkinter import ttk
import openpyxl


def load_data():
    path = "lost_property_book.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)

    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)


def insert_row():
    valuable_property = valuable_property_entry.get()
    hsk_ref = hsk_ref_entry.get()
    ob_number = ob_no_entry.get()
    date_found = date_found_entry.get()
    location_found = location_found_entry.get()
    item_description = item_description_entry.get()
    stored_at = stored_at_entry.get()
    name_of_finder = name_of_finder_entry.get()
    stored_by = stored_by_entry.get()
    remarks = remarks_entry.get()
    dm_name = dm_name_entry.get()
    date_safe_opened = date_safe_opened_entry.get()

    # Insert row into Excel sheet
    path = "lost_property_book.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    row_values = [valuable_property, hsk_ref, ob_number, date_found, location_found, item_description,
                  stored_at, name_of_finder, stored_by, remarks, dm_name, date_safe_opened]

    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('', tk.END, values=row_values)
    
    # Clear the values
    valuable_property_entry.delete(0, "end")
    valuable_property_entry.insert(0, "Valuable No")

    hsk_ref_entry.delete(0, "end")
    hsk_ref_entry.insert(0, "HSK Ref")

    ob_no_entry.delete(0, "end")
    ob_no_entry.insert(0, "OB No")

    date_found_entry.delete(0, "end")
    date_found_entry.insert(0, "Date Found")

    location_found_entry.delete(0, "end")
    location_found_entry.insert(0, "Location Found")

    item_description_entry.delete(0, "end")
    item_description_entry.insert(0, "Item Description")

    stored_at_entry.delete(0, "end")
    stored_at_entry.insert(0, "Stored At")

    name_of_finder_entry.delete(0, "end")
    name_of_finder_entry.insert(0, "Name of Finder")

    stored_by_entry.delete(0, "end")
    stored_by_entry.insert(0, "Stored By")

    remarks_entry.delete(0, "end")
    remarks_entry.insert(0, "Remarks")

    dm_name_entry.delete(0, "end")
    dm_name_entry.insert(0, "DM Name")

    date_safe_opened_entry.delete(0, "end")
    date_safe_opened_entry.insert(0, "Date safe opened")


# creating window
root = tk.Tk()
root.title("GHS Lost Property")
root.resizable(height=False, width=False)

# Creating frame for all the entries
entries_frame = ttk.LabelFrame(text="Insert Information")
entries_frame.grid(row=0, column=0, padx=10, pady=10)

# creating all the entries
valuable_property_entry = ttk.Entry(entries_frame)
valuable_property_entry.insert(0, "Valuable No")
valuable_property_entry.bind("<FocusIn>", lambda e: valuable_property_entry.delete('0', 'end'))
valuable_property_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

hsk_ref_entry = ttk.Entry(entries_frame)
hsk_ref_entry.insert(0, "HSK Ref")
hsk_ref_entry.bind("<FocusIn>", lambda e: hsk_ref_entry.delete('0', 'end'))
hsk_ref_entry.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

ob_no_entry = ttk.Entry(entries_frame)
ob_no_entry.insert(0, "OB No")
ob_no_entry.bind("<FocusIn>", lambda e: ob_no_entry.delete('0', 'end'))
ob_no_entry.grid(row=2, column=0, padx=5, pady=(0, 5), sticky="ew")

date_found_entry = ttk.Entry(entries_frame)
date_found_entry.insert(0, "Date Found")
date_found_entry.bind("<FocusIn>", lambda e: date_found_entry.delete('0', 'end'))
date_found_entry.grid(row=3, column=0, padx=5, pady=(0, 5), sticky="ew")

location_found_entry = ttk.Entry(entries_frame)
location_found_entry.insert(0, "Location Found")
location_found_entry.bind("<FocusIn>", lambda e: location_found_entry.delete('0', 'end'))
location_found_entry.grid(row=4, column=0, padx=5, pady=(0, 5), sticky="ew")

item_description_entry = ttk.Entry(entries_frame)
item_description_entry.insert(0, "Item Description")
item_description_entry.bind("<FocusIn>", lambda e: item_description_entry.delete('0', 'end'))
item_description_entry.grid(row=5, column=0, padx=5, pady=(0, 5), sticky="ew")

stored_at_entry = ttk.Entry(entries_frame)
stored_at_entry.insert(0, "Stored At")
stored_at_entry.bind("<FocusIn>", lambda e: stored_at_entry.delete('0', 'end'))
stored_at_entry.grid(row=6, column=0, padx=5, pady=(0, 5), sticky="ew")

name_of_finder_entry = ttk.Entry(entries_frame)
name_of_finder_entry.insert(0, "Name of Finder")
name_of_finder_entry.bind("<FocusIn>", lambda e: name_of_finder_entry.delete('0', 'end'))
name_of_finder_entry.grid(row=7, column=0, padx=5, pady=(0, 5), sticky="ew")

stored_by_entry = ttk.Entry(entries_frame)
stored_by_entry.insert(0, "Stored By")
stored_by_entry.bind("<FocusIn>", lambda e: stored_by_entry.delete('0', 'end'))
stored_by_entry.grid(row=8, column=0, padx=5, pady=(0, 5), sticky="ew")

remarks_entry = ttk.Entry(entries_frame)
remarks_entry.insert(0, "Remarks")
remarks_entry.bind("<FocusIn>", lambda e: remarks_entry.delete('0', 'end'))
remarks_entry.grid(row=9, column=0, padx=5, pady=(0, 5), sticky="ew")

dm_name_entry = ttk.Entry(entries_frame)
dm_name_entry.insert(0, "DM Name")
dm_name_entry.bind("<FocusIn>", lambda e: dm_name_entry.delete('0', 'end'))
dm_name_entry.grid(row=10, column=0, padx=5, pady=(0, 5), sticky="ew")

date_safe_opened_entry = ttk.Entry(entries_frame)
date_safe_opened_entry.insert(0, "Date safe opened")
date_safe_opened_entry.bind("<FocusIn>", lambda e: date_safe_opened_entry.delete('0', 'end'))
date_safe_opened_entry.grid(row=11, column=0, padx=5, pady=(0, 5), sticky="ew")

# Creating the insert data button
button = ttk.Button(entries_frame, text="Insert", command=insert_row)
button.grid(row=12, column=0, padx=5, pady=5, sticky="nsew")

# Creating frame for the treeview table
treeFrame = ttk.Frame()
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Valuable No", "HSK Ref", "OB No", "Date Found", "Location Found", "Item Description", "Stored At",
        "Name of Finder", "Stored By", "Remarks", "DM Name", "Date safe opened")

treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=17)

treeview.column("Valuable No", width=80)
treeview.column("HSK Ref", width=50)
treeview.column("OB No", width=50)
treeview.column("Date Found", width=75)
treeview.column("Location Found", width=100)
treeview.column("Item Description", width=100)
treeview.column("Stored At", width=70)
treeview.column("Name of Finder", width=100)
treeview.column("Stored By", width=100)
treeview.column("Remarks", width=100)
treeview.column("DM Name", width=100)
treeview.column("Date safe opened", width=110)

treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()

root.mainloop()
