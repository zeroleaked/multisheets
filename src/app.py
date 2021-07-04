from tkinter import *
from tkinter import filedialog
from tabulate import tabulate
import json
import os
from smtplib import SMTPSenderRefused
from time import sleep

from datetime import date

from parser import Spreadsheet
from mailer import Mailer
from utils import read_config, save_config

THRESHOLD = 70000

def bulk_send_mail(mailer, extracts, month, year):

    def print_progress(count, length, number):
        print(f"\r{count}/{length} terkirim, terakhir {number} ", end="")

    length = len(extracts)
    count = 0
    print(f"0/{length} terkirim", end="")

    for row in extracts:
        real_email = row[1]
        mock_email = "fakegags+" + row[1].replace("angkasapura2.co.id", "gmail.com")
        mailer.send_mail(row[0], mock_email, row[2], row[3], month, year)
        count += 1
        print_progress(count, length, row[0])
        sleep(1)

    print()
    
# get path ke file tagihan
def get_table_from_user():
    print("Pilih data tagihan (*.csv / *.xls)")
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    last_dir = read_config("last_data_dir")
    if os.path.exists(last_dir):
        path_ss = filedialog.askopenfilename(initialdir=last_dir)
    else:
        path_ss = filedialog.askopenfilename()
    
    # save path
    table_dir = os.path.dirname(path_ss)
    save_config("last_data_dir", table_dir)

    print("Data tagihan :", path_ss)
    spreadsheet = Spreadsheet(path_ss)
    return spreadsheet

# get indexes
def get_table_config_from_user(table):

    peek = []
    for i in range( len(table[0]) ):
        peek.append([i, table[0][i], table[1][i]])

    print("\n", tabulate(peek, headers=["Indeks", "Judul di Tabel", "Cuplikan Data"] ), "\n")

    number_index = int(input("Indeks kolom nomor telepon : "))
    subtotal_index = int(input("Indeks kolom nomor subtotal (sebelum PPN) : "))

    data_indexes_str = input("Indeks kolom data (dipisah spasi) : ")
    data_indexes_str = data_indexes_str.split()
    print("Label kolom data : ")
    
    labels = []
    for i in range( len(data_indexes_str) ):
        index = int(data_indexes_str[i])
        label = input(f"\t{index} ({table[0][index]}) : ")
        if not label:
            label = table[0][index]
        labels.append(label)

    data_indexes = []
    for i in range( len(labels) ):
        data_indexes.append({
            "index" : int(data_indexes_str[i]),
            "label" : labels[i]
        })

    return number_index, subtotal_index, data_indexes

def get_table_config_from_mem(columns):
    table_configs = read_config("tables")
    for table_config in table_configs:
        if columns == table_config["columns"]:
            return table_config["number_index"], table_config["subtotal_index"], table_config["data_indexes"]
    return None

def save_table_config(columns, number_index, subtotal_index, data_indexes):
    table_config = read_config("tables")
    table_config.append({
        "columns" : columns,
        "number_index" : number_index,
        "subtotal_index" : subtotal_index,
        "data_indexes" : data_indexes
    })
    save_config("tables", table_config)

def get_contacts_path_from_mem():
    contacts_path = read_config("contacts")
    if os.path.exists(contacts_path):
        return contacts_path
    else:
         return None

def get_contacts_path_from_user():
    print("Pilih data email (*.xlsx)")
    contacts_path = filedialog.askopenfilename()
    save_config("contacts", contacts_path)
    return contacts_path
    
def get_month():
    month_name = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    
    today = date.today()
    month = input(f"Bulan tagihan ({today.month}) : ")
    year = input(f"Tahun tagihan ({today.year}) : ")
    if not month:
        month = today.month
    if not year:
        year = today.year

    month = month_name[ int(month) - 1 ]
    return month, year

if __name__ == "__main__":
    spreadsheet = get_table_from_user()

    indexes = get_table_config_from_mem(spreadsheet.table[0])
    if indexes:
        number_index, subtotal_index, data_indexes = indexes
    else:
        number_index, subtotal_index, data_indexes = get_table_config_from_user(spreadsheet.table[:2])
        save_table_config(spreadsheet.table[0], number_index, subtotal_index, data_indexes)

    spreadsheet.define_indexes(number_index, subtotal_index, data_indexes)

    spreadsheet.threshold(THRESHOLD)

    contacts_path = get_contacts_path_from_mem()
    if not contacts_path:
        contacts_path = get_contacts_path_from_user()
    spreadsheet.match_mails(contacts_path)

    extracts = spreadsheet.extract()

    print("\nSample", extracts[0][0])
    for sample in extracts[0][3]:
        print(f"{sample[0]:20s}: {sample[1]}")
    print()

    confirm = input(f"Kirim {len(extracts)} pesan? (Y/N): ")
    if confirm.lower() != "y":
        exit()
    print()

    month, year = get_month()

    mailer = Mailer()

    bulk_send_mail(mailer, extracts, month, year)

    mailer.smtpQuit()

