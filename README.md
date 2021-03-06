# multisheets
A python spreadsheet parser to serialize multiple formats of bill into a database of columns of 1) telephone number and 2) monthly bill total after 10% VAT.
Semiautomated and requires mapping by user input for every **new** table formats. Inputted table format mappings are saved at config.json.

Accepts xml, html, xls, xlsx and csv with ',', '|' and ';' delimiter.

Uploads to MySQL database in localhost with user "root" and password "", can be changed in src/database.py.

## Installation
Use pip to install requirements. List of dependencies can be accessed in requirements.txt.

    pip install -r requirements.txt

For Windows OS, requires additional installation outside of requirements.txt for MySQL Python connector that can be downloaded [here](https://dev.mysql.com/downloads/connector/python/).

## Execution
Open topmost directory in terminal and run src/database.py with Python.

    python src/database.py
