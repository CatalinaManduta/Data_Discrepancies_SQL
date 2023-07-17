# Data_Discrepancies_SQL
This Python program is designed to validate and compare data between a SQL database and an Excel spreadsheet, logging any discrepancies for further examination.

Description
This program reads data from an Excel spreadsheet and compares it row-by-row against data in a specified SQL database table. It's particularly useful for quality assurance tasks, data integrity checks, and for spotting discrepancies that might indicate issues in data entry or transfer processes.

The program examines the action required for each row, specified by the ACTION field in the Excel sheet. This can be one of the following:

"Create"
"Change"
"Delete"
Depending on the action type, the program will search for a matching row in the SQL database and compare each field in the Excel file with the corresponding field in the SQL database. Any discrepancies between Excel and SQL data are logged and saved in separate Excel files for further examination.

Features
Validation of data integrity between Excel and SQL database.
Easy identification of discrepancies with outputs saved in separate Excel files.
Customizable with different actions: "Create", "Change", and "Delete".
Usage
This program is intended to be used in a Python environment where the pandas and pyodbc packages are installed. The Excel file and SQL database should be accessible from this environment.

Please ensure to update the database connection details, file path, and relevant column names as per your specific setup before running the script.

Author: Catalina Manduta
