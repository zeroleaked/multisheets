import mysql.connector

from multisheet_spreadsheet import Spreadsheet
import interface

if __name__ == "__main__":

    path_ss = interface.get_spreadsheet_path_from_user()
    spreadsheet = Spreadsheet(path_ss)

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
    

    #add to database
    cnx = mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="telepon"
    )
    cursor = cnx.cursor()
    insert_stmt = (
        "INSERT INTO tagihan (notel, bulan, total) " +
        "VALUES (%s, %s, %s);"
    )
    count = 0
    for i, row in spreadsheet.minimal_df.iterrows():
        if row["notel"] and row["total"]:
            cursor.execute(insert_stmt, (row["notel"], date, row["total"]))
            count += 1
    cnx.commit()
    print(count, "baris telah ditambah ke database")