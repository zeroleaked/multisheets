# multisheets
A python spreadsheet parser to serialize multiple formats of bill into a database of columns of 1) telephone number and 2) monthly bill total after 10% VAT. 

Accepts xml, html, xls, xlsx and csv with ',', '|' and ';' delimiter.

Uploads to MySQL database in localhost with user "root" and password "", can be changed in src/database.py.

## Installation
For Windows OS, requires additional installation outside of requirements.txt for MySQL Python connector that can be downloaded [here](https://dev.mysql.com/downloads/connector/python/).

## Execution
python src/database.py
