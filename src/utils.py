# {
#     "contacts" : xlsx_path,
#     "last_monthly_dir" : last source dir,
#     "tables" : [
#         {
#             "columns" : [column_name],
#             "number_index" : index,
#             "subtotal_index" : index,
#             "data_indexes" : [
#                 {
#                     "index" : index,
#                     "label" : label
#                 }
#             ]
#         }
#     ]
# }

import json                                                                                                                                                                                                                                                                                                                                                             
import os

def create_config():
    config = {
        "contacts" : False,
        "last_monthly_dir" : False,
        "tables" : []
        }
    with open('config.json', 'w') as outfile:
        json.dump(config, outfile, indent=4)

def read_config(label):
    if not os.path.exists('config.json'):
        create_config()
    with open("config.json") as infile:
        config = json.load(infile)
    return config[label]   

def save_config(label, data):
    if not os.path.exists('config.json'):
        create_config()
    with open('config.json', 'r') as infile:
        config = json.load(infile)
    config[label] = data
    with open('config.json', 'w') as outfile:
        json.dump(config, outfile, indent=4)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
    
# if __name__ == "__main__":
    # path = read_config("contacts")
    # print(path)