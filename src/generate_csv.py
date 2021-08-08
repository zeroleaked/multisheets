import interface

if __name__ == "__main__":
    spreadsheet = interface.get_spreadsheet_from_user()

    indexes = interface.get_table_config_from_mem(spreadsheet.table[0])
    if indexes:
        number_index, subtotal_index, data_indexes = indexes
    else:
        number_index, subtotal_index, data_indexes = interface.get_table_config_from_user(spreadsheet.table[0], spreadsheet.table[1])
        interface.save_table_config(spreadsheet.table[0], number_index, subtotal_index, data_indexes)

    spreadsheet.define_indexes(number_index, subtotal_index, data_indexes)

    spreadsheet.calculate_total()

    month, year = interface.get_month('int')
    month_str = str(month)
    if len(month_str) < 2:
        month_str = '0' + month_str
    date = str(year) + '-' + month_str + '-' + "25"
    
    spreadsheet.export_csv(date)