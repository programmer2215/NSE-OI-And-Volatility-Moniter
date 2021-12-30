import requests as req
from datetime import datetime
import dload
import csv
import os
import tkinter as tk
from tkinter import ttk
import tkcalendar as tkcal 

# Last Version With Vola

root = tk.Tk()
root.title("Futures Volume & OI Moniter")

style = ttk.Style()
style.configure("Treeview", font=('Britannic', 11, 'bold'), rowheight=25)
style.configure("Treeview.Heading", font=('Britannic' ,13, 'bold'))

# Tkinter Bug Work Around
if root.getvar('tk_patchLevel')=='8.6.9': #and OS_Name=='nt':
    def fixed_map(option):
        # Fix for setting text colour for Tkinter 8.6.9
        # From: https://core.tcl.tk/tk/info/509cafafae
        #
        # Returns the style map for 'option' with any styles starting with
        # ('!disabled', '!selected', ...) filtered out.
        #
        # style.map() returns an empty list for missing options, so this
        # should be future-safe.
        return [elm for elm in style.map('Treeview', query_opt=option) if elm[:2] != ('!disabled', '!selected')]
    style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))

frame_top = tk.Frame(root)
frame_top.pack(padx=5, pady=20)

tv = ttk.Treeview(
    frame_top, 
    columns=(1, 2, 3, 4), 
    show='headings', 
    height=10) # add column 4 when you need to Add Volatility Back
tv.pack()

tv.heading(1, text='Security')
tv.heading(2, text='OI Cng')
tv.heading(3, text='Volume')
tv.heading(4, text='Price Cng %')
#tv.heading(4, text='Volatility')

def delete_data_files(DATA_DIRECTORY):
    datafilelist = [ f for f in os.listdir(DATA_DIRECTORY) if not f == "stocks.txt"]
    for file in datafilelist:
        os.remove(os.path.join(DATA_DIRECTORY,file))

def export_csv():
    EXPORT_FILE = f"C:\\Users\\Administrator\\Desktop\\csv files\\FUTURES\\SortedData.csv"
    with open(EXPORT_FILE, 'w', newline='') as f:
        csv_writer = csv.writer(f, delimiter=",")
        for line in tv.get_children():
            csv_writer.writerow([tv.item(line)['values'][0]])
            
            

def calc():
    day = calender.get_date()
    is_nifty_50 = var1.get()
    print(is_nifty_50)
    day_str = day.strftime('%d%m%Y')

    #formatting for volume URL
    month = day.strftime('%b').upper()
    year = day.strftime('%Y')
    date_volume = day.strftime('%d%b%Y').upper()

    DATA_DIRECTORY = ".\\Data Files\\"
    VOL_OI_URL = f"https://www1.nseindia.com/content/historical/DERIVATIVES/{year}/{month}/fo{date_volume}bhav.csv.zip"

    #VOLATILITY_URL = f"https://www1.nseindia.com/archives/nsccl/volt/CMVOLT_{day_str}.CSV"
    print(VOL_OI_URL)
    dload.save_unzip(VOL_OI_URL, f"{DATA_DIRECTORY}", delete_after=True)

    volume_file_name = f"fo{date_volume}bhav.csv"
    FULL_PATH_VOLUME_OI = os.path.join(DATA_DIRECTORY, volume_file_name)

    stocks_data = {}

    if is_nifty_50 == 1:
        with open("stocks.txt", "r") as f:
            stocks = [stock.strip() for stock in f]
    else:
        with open(FULL_PATH_VOLUME_OI) as f:
            csv_reader = csv.reader(f, delimiter=",")
            next(csv_reader, None)
            stocks = []
            for row in csv_reader:
                if row[1] not in stocks:
                    stocks.append(row[1])
    

    with open(FULL_PATH_VOLUME_OI) as f:
        csv_reader = csv.reader(f, delimiter=",")

        for row in csv_reader:
            if row[1] in stocks and row[0] == 'FUTSTK':
                price_change = round(float(row[8]) - float(row[5]), 2)
                if int(row[12]) - int(row[13]) != 0:
                    oi_change_per = round((int(row[13]) / (int(row[12]) - int(row[13])))*100, 2)
                else:
                    oi_change_per = 0.0
                if price_change != 0 and float(row[5]) != 0:
                    price_change_per = round(((price_change / float(row[5])) * 100), 2)
                else:
                    price_change_per = 0.0
                    
                if row[1] not in stocks_data:
                    stocks_data[row[1]] = [[oi_change_per], int(row[10]), price_change_per, row[2]]
                else:
                    stocks_data[row[1]][0].append(oi_change_per)
                    stocks_data[row[1]][1] += int(row[10])
    for i in stocks_data.keys():
        stocks_data[i][0] =  max(stocks_data[i][0])

    #delete_data_files(".\\Data Files\\")
    return stocks_data

controls_frame = tk.Frame(root)
controls_frame.pack()

calender_lab = tk.Label(controls_frame, text="Date:")
calender_lab.grid(row=0, column=0, rowspan=2, pady=5, padx=10)
calender = tkcal.DateEntry(controls_frame, selectmode="day")
calender.grid(row=1, column=0, rowspan=2, pady=5, padx=10)

row_lab = ttk.Label(controls_frame, text="Show Rows:")
row_lab.grid(row=0, column=1, pady=5, rowspan=2, padx=10)
row_entry_var = tk.StringVar(value="10")
row_entry = ttk.Entry(controls_frame, textvariable=row_entry_var)
row_entry.grid(row=1, column=1, rowspan=2, pady=5, padx=10)

#print(nifty_50_data)

def set_table():
    for i in tv.get_children():
        tv.delete(i)
    stocks_data = calc()
    i = 0
    no_of_rows = int(row_entry_var.get())
    default_sort_order = default_selected.get()
    
    if default_sort_order == "vol":
        default_sort_index = 1
    elif default_sort_order == "OI":
        default_sort_index = 0
    
    data_vol = sorted(stocks_data.items(), key=lambda x: x[1][default_sort_index], reverse=True)[:no_of_rows]
    
    sort_order = selected.get()

    if sort_order == "vol":
        sort_index = 1
    elif sort_order == "OI":
        sort_index = 0

    for k,v in sorted(data_vol, key=lambda x: x[1][sort_index], reverse=True):
        tv.insert(parent='', index=i, iid=i, values=(k, v[0], v[1], v[2]))
        i += 1
    


selected = tk.StringVar(value="vol")
filter_lab = tk.Label(controls_frame, text="---FILTER---").grid(row=0, column=2, pady=5, padx=5)
r1 = ttk.Radiobutton(controls_frame, text='Open Interest', value='OI', variable=selected, command=set_table)
r1.grid(row=1, column=2, pady=5, padx=10, sticky=tk.W)
r2 = ttk.Radiobutton(controls_frame, text='Volume', value='vol', variable=selected, command=set_table)
r2.grid(row=2, column=2, pady=5, padx=10, sticky=tk.W)


filter_lab = tk.Label(controls_frame, text="---DEFAULT FILTER---").grid(row=0, column=3, pady=5, padx=10)

default_selected = tk.StringVar(value="OI")
r3 = ttk.Radiobutton(controls_frame, text='Open Interest', value='OI', variable=default_selected, command=set_table)
r3.grid(row=1, column=3, pady=5, padx=10, sticky=tk.W)
r4 = ttk.Radiobutton(controls_frame, text='Volume', value='vol', variable=default_selected, command=set_table)
r4.grid(row=2, column=3, pady=5, padx=10, sticky=tk.W)

var1 = tk.IntVar(value="0")
nifty_50_check = ttk.Checkbutton(root, text="Nifty 50", variable=var1)
nifty_50_check.pack(pady=15)

search_button = ttk.Button(root, text="Search", command=set_table)
search_button.pack(pady=20)

export_button = ttk.Button(root, text="Export", command=export_csv)
export_button.pack()

root.mainloop()


