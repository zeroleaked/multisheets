from tkinter import *
from tkinter import filedialog
from tabulate import tabulate
import os
from time import sleep
import getpass
from datetime import date

from multisheet_spreadsheet import Spreadsheet
from mailer import Mailer
from utils import read_config, save_config

THRESHOLD = 60000

# bulk_send_mail(mailer, extracts, month, year)
# send mails by information from extracts, returns None
# params:
#   mailer      : Mailer object
#   extracts    : list of dict
#                   [{
#                       "email": String,
#                       "jabatan": String,
#                       "data": {
#                           "Nomor Telepon": String,
#                           "Unit": String,
#                           ... optional keys ...,
#                           "Subtotal": String,
#                           "PPN": String,
#                           "Total": String
#                       }
#                   }]
#   month       : int or String as month of bill
#   year        : int as year of bill
def bulk_send_mail(mailer, extracts, month, year):

    length = len(extracts)
    count = 0
    print(f"0/{length} terkirim", end="")

    for extr in extracts:
        
        real_email = extr["email"]
        # real_email = "maryam.namira@angkasapura2.co.id"
        mock_email = "fakegags+" + real_email.lower().replace("angkasapura2.co.id", "gmail.com")
        extr["email"] = mock_email

        mailer.send_mail(extr, month, year)
        count += 1
        print(f"\r{count}/{length} terkirim, terakhir {extr['email']}              ", end="")

        sleep(1)

    print()
    
# get_spreadsheet_from_user()
# get path to monthly bill file from user, returns Spreadsheet object.
def get_spreadsheet_from_user():
    print("Pilih data tagihan")

    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    last_dir = read_config("last_monthly_dir")
    if os.path.exists(last_dir):
        path_ss = filedialog.askopenfilename(initialdir=last_dir)
    else:
        path_ss = filedialog.askopenfilename()
    
    # save path for next app run
    table_dir = os.path.dirname(path_ss)
    save_config("last_monthly_dir", table_dir)

    print("Data tagihan :", path_ss)
    spreadsheet = Spreadsheet(path_ss)
    return spreadsheet

# get_table_config_from_user(headers, sample)
# params:
#   headers : list of column labels in String, 
#   sample  : list of a sample row in String
# get specific indexes of columns from user. Returns telephone number column index,
# subtotal column index, and optional columns indexes from user
def get_table_config_from_user(headers, sample):

    # show all column labels and first row of data
    peek = []
    for i in range( len(headers) ):
        peek.append([i, headers[i], sample[i]])
    print("\n", tabulate(peek, headers=["Indeks", "Judul di Tabel", "Cuplikan Data"] ), "\n")

    number_index = int(input("Indeks kolom nomor telepon : "))
    subtotal_index = int(input("Indeks kolom nomor subtotal (sebelum PPN) : "))

    data_indexes_str = input("Indeks kolom data opsional (dipisah spasi) : ")
    data_indexes_str = data_indexes_str.split()
    print("Label kolom data : ")
    labels = []
    for i in range( len(data_indexes_str) ):
        index = int(data_indexes_str[i])
        label = input(f"\t{index} ({headers[index]}) : ")
        if not label:
            label = headers[index]
        labels.append(label)

    data_indexes = []
    for i in range( len(labels) ):
        data_indexes.append({
            "index" : int(data_indexes_str[i]),
            "label" : labels[i]
        })

    return number_index, subtotal_index, data_indexes

# get_table_config_from_mem(columns):
# params:
#   columns: list of column labels as String
# compares given columns with saved columns in config.json. Returns telephone number column index,
# subtotal column index, and optional columns indexes if exists. Returns None if config.json not
# found or columns not found in config.json.
def get_table_config_from_mem(columns):
    table_configs = read_config("tables")
    for table_config in table_configs:
        if columns == table_config["columns"]:
            return table_config["number_index"], table_config["subtotal_index"], table_config["data_indexes"]
    return None

# save_table_config(columns, number_index, subtotal_index, data_indexes):
# params:
#   columns         : list of String as column labels
#   number_index    : int as telephone number column index
#   subtotal_index  : int as subtotal column index
#   data_indexes    : list of dict {index, label} with index as opritonal column index and label
#                       as its column label replacement
# saves column indexes configuration to config.json for the specified list of column labels, returns None
def save_table_config(columns, number_index, subtotal_index, data_indexes):
    table_config = read_config("tables")
    table_config.append({
        "columns" : columns,
        "number_index" : number_index,
        "subtotal_index" : subtotal_index,
        "data_indexes" : data_indexes
    })
    save_config("tables", table_config)

# get_contacts_path_from_mem():
# reads path to contact file *.xlsx from config.json.
# Returns String as path if path exists, returns None if path does not exist. 
def get_contacts_path_from_mem():
    contacts_path = read_config("contacts")
    if os.path.exists(contacts_path):
        return contacts_path
    else:
         return None

# get_contacts_path_from_user():
# pops a filedialog for user to input path to contact file *.xlsx.
# Returns String as path 
def get_contacts_path_from_user():
    print("Pilih data email (*.xlsx)")
    contacts_path = filedialog.askopenfilename()
    save_config("contacts", contacts_path)
    return contacts_path

# print_extracts(extracts):
# prints contents of extracts to terminal. See bulk_send_mail for extracts format
def print_extracts(extracts):
    for ex in extracts:    
        print("\n", ex["jabatan"], ex["email"])
        headers = list(ex["data"][0].keys())
        peek = []
        for row in ex["data"]:
            peek_row = []
            for key in row:
                peek_row.append(row[key])
            peek.append(peek_row)
        print(tabulate(peek, headers=headers, floatfmt=".3f"), "\n")
    print()
    
# get_month():
# get month and year from user
def get_month(type):
    month_name = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    
    today = date.today()
    month = input(f"Bulan tagihan ({today.month - 1}) : ")
    year = input(f"Tahun tagihan ({today.year}) : ")
    if not month:
        month = today.month - 1
    if not year:
        year = today.year

    if type == 'str':
        month = month_name[ int(month) - 1 ]
    else:
        month = int(month)
        year = int(year)

    return month, year

def confirm(length):
    ans = input(f"Kirim {length} pesan? (Y/N): ")
    if ans.lower() != "y":
        exit()

def email_routine():
    spreadsheet.threshold(THRESHOLD)

    contacts_path = get_contacts_path_from_mem()
    if not contacts_path:
        contacts_path = get_contacts_path_from_user()
    spreadsheet.match_mails(contacts_path)

    extracts = spreadsheet.extract_dict()
    print_extracts(extracts)
    
    confirm(len(extracts))

    month, year = get_month('string')

    sender_email = input("Email pengirim: ")
    sender_password = getpass.getpass()

    mailer = Mailer(sender_email, sender_password)

    bulk_send_mail(mailer, extracts, month, year)

def export_routine():
    month, year = get_month('int')
    month_str = str(month)
    if len(month_str) < 2:
        month_str = '0' + month_str
    date = str(year) + '-' + month_str + '-' + "25"
    spreadsheet.export_csv(date)


if __name__ == "__main__":
    spreadsheet = get_spreadsheet_from_user()

    indexes = get_table_config_from_mem(spreadsheet.table[0])
    if indexes:
        number_index, subtotal_index, data_indexes = indexes
    else:
        number_index, subtotal_index, data_indexes = get_table_config_from_user(spreadsheet.table[0], spreadsheet.table[1])
        save_table_config(spreadsheet.table[0], number_index, subtotal_index, data_indexes)

    spreadsheet.define_indexes(number_index, subtotal_index, data_indexes)

    spreadsheet.calculate_total()

    print()
    print("1 - Ekspor ke CSV")
    print("2 - Email")
    op = input("Pilih : ")
    if op == "1":
        print()
        export_routine()
    elif op == "2":
        print()
        email_routine()
