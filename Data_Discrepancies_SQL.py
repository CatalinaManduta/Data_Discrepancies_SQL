# Import data
import snoop
from openpyxl import load_workbook
import pandas as pd
import pyodbc

wb = load_workbook('Book1_test.xlsx')
ab = wb.properties

metadata_dict = vars(ab)
metadata_df = pd.DataFrame.from_dict(metadata_dict, dtype="str", orient="index")
# print(metadata_df)
last_modif = metadata_df.loc['lastModifiedBy']
print(last_modif)
creator = metadata_df.loc['creator']
print(creator)



cursor = conn.cursor()

def connect_to_server():
    """
      This function establishes a connection to a specific SQL Server database,
      using a trusted connection method.

      It connects to the database named 'ServerName' located on the server 'DataBase'
      using the 'SQL SERVER' driver. It utilizes Windows authentication
      (as specified by 'Trusted_Connection=yes') which uses the current user's login credentials to authenticate.

      Once the connection is established, a cursor object is created. The cursor acts as a
      medium for executing SQL statements and fetching results.

      Returns:
          cursor: An instance of the Cursor class that has been instantiated from the Connection object.
                  This object can be used to execute SQL commands.

      Note:
          This function doesn't handle exceptions. In production code, consider handling potential
          errors such as failed connection or unsuccessful cursor creation.
      """
    conn = pyodbc.connect('Driver={SQL SERVER};'
                          'Server=ServerName;'
                          'Database=DataBase;'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()


xls = pd.ExcelFile('Book1_test.xlsx')
df = xls.parse('Request', skiprows=7)


# Clean the data
def clean(dataframe, column: str):
    """
    This function cleans the specified column in the provided dataframe
    by replacing NaN (Not a Number) values with an empty string.

    It uses the pandas DataFrame `fillna` method, which fills NA/NaN values
    using the specified method. Here, we're replacing NaNs with an empty string ('').
    The 'inplace=True' argument tells pandas to modify the existing DataFrame directly,
    instead of creating a new, modified DataFrame.

    Parameters:
    dataframe (pandas.DataFrame): The DataFrame to clean.
    column (str): The column in the DataFrame that should be cleaned.

    Returns:
    None. The function modifies the DataFrame in-place.
    """
    dataframe[column].fillna('', inplace=True)


clean(df, 'Column1')
clean(df, 'Column2')


# Create an empty dataframe where the rows which have errors will be saved
error_data = {
    'ERROR LINE': [],
    'Action': [],
    'ID': [],
    'Column4': [],
    'Column5': [],
    'Column6': [],
    'Column7': [],
    'Column8': [],
    'Column9': [],
    'Column10': [],
    'Column11': [],
    'Column12': [],
    'Column13': [],
    'Column14': [],
    'Column15': [],
    'Column16': [],
    'ERROR MESSAGE': []}

error_data = pd.DataFrame(error_data)
print(error_data)
extra_info = pd.DataFrame({'Last MODIFIER': last_modif,
                          'Last CREATOR': creator})


def check_ID(ID):  # ID here is the ID number which must be checked in the SQL database
    """
    This function checks if a given ID is present in the 'ID' column
    of the '[dbo].[Forecast]' table in a connected SQL database.

    It uses the cursor's execute method to run a SELECT SQL query on the
    database, which fetches all records where 'ID' matches the provided ID.

    Parameters:
    ID: The ID value that needs to be checked for existence in the database.
        The type of ID should match the data type of 'ID' in the database.

    Returns:
    bool: True if the ID exists in the database (i.e., the SELECT statement returns
          at least one record). False otherwise (the SELECT statement returns no records).

    Example:
    >>> check_ID(1234)
    True

    Note:
    It's assumed that the 'cursor' object is already created before this function is called.
    'cursor' should be a valid pyodbc.Cursor instance connected to the target SQL database.
    """
    var = cursor.execute('''SELECT ID FROM [dbo].[Forecast]  WHERE ID = ?''', ID)
    row = cursor.fetchall()
    if len(row)>0:
        return True
    else:
        return False


# Complete version
connect_to_server()  # this function wll connect you automatically tot the server


def update_and_errors(dataframe): # the argument here is the dataframe where you want your error rows to be saved
    """
    This function iterates over each row of the provided dataframe (derived from an excel file),
    performs actions on a SQL database based on the 'ACTION' column in the dataframe,
    and saves any errors that occur during this process in a new excel file.

    The function assumes that 'df' is a globally defined pandas DataFrame
    that is being referred inside the function.

    The actions can be 'Create', 'Delete' and 'Change'. For each action, different SQL
    operations are performed (Insert, Delete, and Update respectively).

    In case an error occurs during any action, the function will append an error message
    to the dataframe along with the information about the row where the error occurred.

    Parameters:
    dataframe (pandas.DataFrame): The DataFrame where the error rows are saved.
    Returns:
    dataframe (pandas.DataFrame): The DataFrame containing the error rows.
                                  The DataFrame is also saved as 'Final_error_file.xlsx'.

    Note:
    The function assumes a valid pyodbc.Cursor instance (cursor) is available in the global scope.
    The cursor should be connected to the target SQL database.
    The 'conn' is a pyodbc.Connection object.
    """
    iteration = 1
    for i in range(len(df)):
        try:
            # rename each column so we shorten the code
            Action = df.loc[i, 'Action']
            ID = df.loc[i, 'ID']
            Column4 = df.loc[i, 'Column4']
            Column5 = df.loc[i, 'Column5']
            Column6 = df.loc[i, 'Column6']
            Column7 = df.loc[i, 'Column7']
            Column8 = int(df.loc[i, 'Column8'])
            Column9 = df.loc[i, 'Column9']
            Column10 = df.loc[i, 'Column10']
            Column11 = df.loc[i, 'Column11']
            Column12 = df.loc[i, 'Column12']
            Column13 = df.loc[i, 'Column13']
            Column14 = df.loc[i, 'Column14']
            Column15 = int(df.loc[i, 'Column15'])
            value = Action, ID, Column4, Column5, Column6, Column7, Column8, Column9, Column10, \
                    str(Column11), str(Column12), Column13, Column14, Column15

            if Action == "Create":
                if ID == isinstance(ID, int):
                    if ID >= 0 or ID < 0:
                        value_not_found = 'The ID should be empty instead it contains', ID
                        value = [int(iteration), Action, ID, Column4, Column5, Column6, Column7, Column8, Column9,
                                 Column10, str(Column11), str(Column12), Column13, Column14, Column15,
                                 value_not_found]
                        row = pd.Series(value, index=dataframe.columns)
                        dataframe = dataframe.append(row, ignore_index=True)
                        iteration += 1
                        print(f'The ID should be empty instead it contains: {ID}')
                elif ID == isinstance(ID, str):
                    print('This is a string')
                    pass

                else:
                    cursor.execute('''
                        INSERT INTO [dbo].[Test] (Column4, Column5, Column6, Column7, Column8, 
                        Column9, Column10, Column11, Column12, Column13, Column14, Column15)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                   (Column4, Column5, Column6, Column7, Column8, Column9,
                                    Column11, Column12, Column13, Column14, Column15))
                    iteration += 1
                    conn.commit()
                    iteration += 1
                    print('Entry CREATED')

            elif Action == "Delete":
                checker = check_ID(ID)
                if checker:
                    cursor.execute('''DELETE FROM [dbo].[Test] WHERE ID = ?''', (ID))
                    conn.commit()
                    iteration += 1
                    print(f'The {ID} was DELETED!')
                else:
                    value_not_found = 'The', ID, 'given was NOT FOUND in our database!'
                    value = [int(iteration), Action, ID, Column4, Column5, Column6, Column7, Column8, Column9,
                             Column10, str(Column11), str(Column12), Column13, Column14, Column15,
                             value_not_found]
                    row = pd.Series(value, index=dataframe.columns)
                    dataframe = dataframe.append(row, ignore_index=True)
                    iteration += 1
                    print(f'The {ID} given was NOT FOUND in our database!')

            elif Action == 'Change':
                checker = check_ID(ID)
                if checker:
                    cursor.execute('''
                        UPDATE [dbo].[Test] 
                        SET Column4 =?, Column5 =?, Column6 =?, Column7 =?, Column8 =?, 
                        Column9 =?, Column10 =?, Column11 =?, Column12 =?, Column13 =?, 
                        Column14 =?, Column15 =? WHERE ID = ?''',
                                   Column4, Column5, Column6, Column7, Column8, Column9,
                                   Column11, Column12, Column13, Column14, Column15, ID)
                    conn.commit()
                    iteration += 1
                    print(f'The {ID} was CHANGED')
                else:
                    value_not_found = 'The', ID, 'given was NOT FOUND in our database!'
                    value = [int(iteration), Action, ID, Column4, Column5, Column6, Column7, Column8, Column9,
                             Column10, str(Column11), str(Column12), Column13, Column14, Column15,
                             value_not_found]
                    row = pd.Series(value, index=dataframe.columns)
                    dataframe = dataframe.append(row, ignore_index=True)
                    iteration += 1
                    print(f'The {ID} given was NOT FOUND in our database!')

        except Exception as e:
            iteration += 1
            value = [int(iteration), Action, ID, Column4, Column5, Column6, Column7, Column8, Column9,
                     Column10, str(Column11), str(Column12), Column13, Column14, Column15, e]
            if Action == "Create":
                print(f'''The error apeared in CREATE section row number {iteration}.''')
                row = pd.Series(value, index=dataframe.columns)
                print(e)
                dataframe = dataframe.append(row, ignore_index=True)
            elif Action == "Delete":
                print(f'''The error apeared in DELETE section row number {iteration}.''')
                row = pd.Series(value, index=dataframe.columns)
                print(e)
                dataframe = dataframe.append(row, ignore_index=True)
            elif Action == 'Change':
                print(f'''The error apeared in CHANGE section row number {iteration}.''')
                row = pd.Series(value, index=dataframe.columns)
                print(e)
                dataframe = dataframe.append(row, ignore_index=True)
            continue
    dataframe = dataframe.append(extra_info)
    dataframe.to_excel("Final_error_file.xlsx")
    return dataframe

update_and_errors(error_data)


def check_row(ID):
    """
        This function retrieves all the columns of a row in the '[dbo].[Test]' table in a
        connected SQL database that matches the provided 'ID'.

        It uses the cursor's execute method to run a SELECT SQL query on the
        database, which fetches the entire row where 'ID' matches the provided ID.
        The row is then converted into a list, except for the last element, which is popped out.

        Parameters:
        ID: The ID value that needs to be checked for existence in the database.
            The type of ID should match the data type of 'ID' in the database.

        Returns:
        list: A list of all the column values from the retrieved row, excluding the last column.
              If the provided ID does not exist in the table, an empty list is returned.
        Note:
        It's assumed that the 'cursor' object is already created before this function is called.
        'cursor' should be a valid pyodbc.Cursor instance connected to the target SQL database.
    """

    var = cursor.execute('''SELECT * FROM [dbo].[Test] WHERE ID = ?''', ID)
    row = cursor.fetchall()
    checking_list = []
    if len(row)>0:
        for i in row[0]:
            checking_list.append(i)
    else:
        return checking_list
    checking_list.pop(-1)
    return checking_list


def check_row_without_condition():
    """
       This function iterates over every row in the dataframe 'df' and for each row,
       it retrieves matching rows from the '[dbo].[Test]' table in a connected SQL database.

       The function creates a list of values for each row in the dataframe and executes
       a SELECT SQL query on the database. If any rows are found in the database that match
       the entire set of column values, those rows are converted into a list and appended to
       the overall 'checking_list'.

       This function assumes that 'df' is a global pandas DataFrame that is being referred inside
       the function.

       Returns:
       list: A list of lists where each sublist contains the values of a row from the database
             that matches the respective row in the dataframe. If no matching rows are found,
             'checking_list' will be an empty list.

       Note:
       It's assumed that the 'cursor' object is already created before this function is called.
       'cursor' should be a valid pyodbc.Cursor instance connected to the target SQL database.
    """
    checking_list = []
    for i in range(len(df)):
        # Iterate over each line in the excel file.
        # Rename each column so we shorten the code.
        Action = df.loc[i, 'Column2']
        ID = df.loc[i, 'Column3']
        Column4 = df.loc[i, 'Column4']
        Column5 = df.loc[i, 'Column5']
        Column6 = df.loc[i, 'Column6']
        Column7 = df.loc[i, 'Column7']
        Column8 = int(df.loc[i, 'Column8'])
        Column9 = df.loc[i, 'Column9']
        Column10 = df.loc[i, 'Column10']
        Column11 = df.loc[i, 'Column11']
        Column12 = df.loc[i, 'Column12']
        Column13 = df.loc[i, 'Column13']
        Column14 = df.loc[i, 'Column14']
        Column15 = df.loc[i, 'Column15']
        Column16 = int(df.loc[i, 'Column16'])

        var = cursor.execute(
            '''SELECT * FROM [dbo].[Test] WHERE Column4 = ? AND Column5 = ? AND Column6 = ? AND Column7 = ? AND Column8 = ? AND Column9 = ? AND Column10 = ? AND Column11 = ? AND Column12 = ? AND Column13 = ? AND Column14 = ? AND Column15 = ? AND Column16 = ?''',
            Column4, Column5, Column6, Column7, Column8, Column9, Column10, Column11, Column12, Column13, Column14,
            Column15, Column16)

        row = cursor.fetchall()
        if len(row) > 0:
            empty = []
            for i in row[0]:
                empty.append(i)
            checking_list.append(empty)
    return checking_list


check_row_without_condition()


unmached_data = {
    'ID': [],
    'Column2': [],
    'Column3': [],
    'Column4': [],
    'Column5': [],
    'Column6' : [],
    'Column7': [],
    'Column8': [],
    'Column9': [],
    'Column10': [],
    'Column11': [],
    'Column12': [],
    'Column13': [],
    'Column14': [],
    'Column15': [],
    'Column16': []}

unmached_data = pd.DataFrame(unmached_data)

unmatched_data = {
    'Column1': [],
    'Column2': [],
    'Column3': [],
    'Column4': [],
    'Column5': [],
    'Column6' : [],
    'Column7': [],
    'Column8': [],
    'Column9': [],
    'Column10': [],
    'Column11': [],
    'Column12': [],
    'Column13': [],
    'Column14': [],
    'Column15': [],
    'Column16': []}


unmached_data2 = pd.DataFrame(unmached_data2)


def check_row_without_condition(Column2, Column3, Column4, Column5, Column6, Column7, Column8, Column9, Column10, Column11, Column12, Column13, Column14, Column15):
    """
    This function receives several parameters representing specific values.
    It then queries the '[dbo].[Test3]' table in a connected SQL database for a row that matches
    all of these parameter values.

    Parameters:
    Column2 (str): The value of the second column.
    Column3 (str): The value of the third column.
    Column4 (str): The value of the fourth column.
    Column5 (str): The value of the fifth column.
    Column6 (int): The value of the sixth column.
    Column7 (str): The value of the seventh column.
    Column8 (str): The value of the eighth column.
    Column9 (date): The value of the ninth column.
    Column10 (date): The value of the tenth column.
    Column11 (str): The value of the eleventh column.
    Column12 (str): The value of the twelfth column.
    Column13 (str): The value of the thirteenth column.
    Column14 (int): The value of the fourteenth column.
    Column15 (int): The value of the fifteenth column.

    Returns:
    list: A list that contains the values of a row from the database that matches the input parameters.
        If no matching row is found, 'checking_list' will be an empty list.
        Note: The dates in the returned list are converted to string format.

    Note:
    It's assumed that the 'cursor' object is already created before this function is called.
    'cursor' should be a valid pyodbc.Cursor instance connected to the target SQL database.
    """
    checking_list = []
    var = cursor.execute('''SELECT Column2 =?, Column3 =?, Column4=?, Column5=?, Column6=?, Column7=?, Column8=?, Column9=?, Column10=?, Column11=?, Column12=?, Column13=?, Column14=?, Column15=? FROM [dbo].[Test3]''', Column2, Column3, Column4, Column5, Column6, Column7, Column8, Column9, Column10, Column11, Column12, Column13, Column14, Column15)
    row = cursor.fetchall()
    if len(row)>0:
        for i in row[0]:
            checking_list.append(i)
    checking_list[7] = str(checking_list[7])
    checking_list[8] = str(checking_list[8])
    return checking_list


#@ snoop
def value_checking(dataframe, dataframe2):
    """
    This function is designed to check for discrepancies between two given DataFrames.
    It iterates over each row in the first DataFrame and checks the corresponding data in the second DataFrame.

    If discrepancies are found, they are logged into the provided DataFrames and these DataFrames are
    written into Excel files for further examination.

    Parameters:
    dataframe (DataFrame): The DataFrame to log discrepancies found.
    dataframe2 (DataFrame): The second DataFrame to check discrepancies against.

    Returns:
    None

    Side Effects:
    Creates Excel files 'Data not properly transfered in SQL: DELETE, CHANGE Test3.xlsx' and
    'Data not properly transfered SQL: CREATE in Test3.xlsx' containing the logged discrepancies.
    """
    iteration = 1  # this helps measure each iteration in order to check when the error appeared
    for i in range(len(df)):
        # iterate over each line in the DataFrame
        try:
            # rename each column so we shorten the code
            Column1 = df.loc[i, 'Column1']
            Column2 = int(df.loc[i, 'Column2'])
            Column3 = df.loc[i, 'Column3']
            Column4 = df.loc[i, 'Column4']
            Column5 = df.loc[i, 'Column5']
            Column6 = int(df.loc[i, 'Column6'])
            Column7 = df.loc[i, 'Column7']
            Column8 = df.loc[i, 'Column8']
            Column9 = df.loc[i, 'Column9']
            Column10 = df.loc[i, 'Column10']
            Column11 = df.loc[i, 'Column11']
            Column12 = df.loc[i, 'Column12']
            Column13 = df.loc[i, 'Column13']
            Column14 = int(df.loc[i, 'Column14'])
            Column15 = int(df.loc[i, 'Column15'])

            df_value = [Column2, Column3, Column4, Column5, Column6, Column7, Column8, str(Column9), str(Column10),
                        Column11, Column12, Column13, Column14, Column15]
            df_value_create = [Column3, Column4, Column5, Column6, Column7, Column8, Column9, str(Column10),
                               str(Column11), Column12, Column13, Column14, Column15]

            if Column1 == "Create":
                SQL_value = check_row_without_condition(Column3, Column4, Column5, Column6, Column7, Column8, Column9,
                                                        Column10, Column11, Column12, Column13, Column14, Column15)

                if len(SQL_value) > 0:
                    for i, j in zip(df_value_create, SQL_value):
                        worked = []
                        if i == j:
                            print(f'Same values: EXCEL:{i}, SQL:{j}.')
                            worked.append(True)
                        else:
                            print(f'DIFFERENT values: EXCEL:{i}, SQL:{j}.')
                            row_excel = pd.Series(df_value_create, index=dataframe2.columns)
                            row_sql = pd.Series(SQL_value, index=dataframe2.columns)
                            row_empty = pd.Series(
                                [f'Iteration:{iteration}', None, None, None, None, None, None, None, None, None, None,
                                 None, None], index=dataframe2.columns)
                            dataframe2 = dataframe2.append(row_empty, ignore_index=True)
                            dataframe2 = dataframe2.append(row_excel, ignore_index=True)
                            dataframe2 = dataframe2.append(row_sql, ignore_index=True)
                    if any(worked):
                        print('The row was successfully CREATED')
                else:
                    print(f'The {Column2} was not found in SQL database')
                    iteration += 1
            elif Column1 == 'Change':
                SQL_value = check_row(Column2)
                if len(SQL_value) > 0:
                    for i, j in zip(df_value, SQL_value):
                        worked = []
                        if i == j:
                            print(f'Same values: EXCEL:{i}, SQL:{j}.')
                            worked.append(True)
                        else:
                            print(f'DIFFERENT values: EXCEL:{i}, SQL:{j}.')
                            row_excel = pd.Series(df_value, index=dataframe.columns)
                            row_sql = pd.Series(SQL_value, index=dataframe.columns)
                            row_empty = pd.Series(
                                [f'Iteration:{iteration}', None, None, None, None, None, None, None, None, None, None,
                                 None, None, None], index=dataframe.columns)
                            dataframe = dataframe.append(row_empty, ignore_index=True)
                            dataframe = dataframe.append(row_excel, ignore_index=True)
                            dataframe = dataframe.append(row_sql, ignore_index=True)
                    if any(worked):
                        print('The row was succesfully CHANGED')
                else:
                    print(f'The {Column2} was not found in SQL database')
                iteration = 1
            elif Column1 == 'DELETE':
                SQL_value = check_row(Column2)
                if len(SQL_value) == 0:
                    print('The values was correctly deleted')
                else:
                    row_excel = pd.Series(df_value, index=dataframe.columns)
                    row_sql = pd.Series(SQL_value, index=dataframe.columns)
                    row_empty = pd.Series(
                        [f'Iteration:{iteration}', None, None, None, None, None, None, None, None, None, None, None,
                         None, None], index=dataframe.columns)
                    dataframe = dataframe.append(row_empty, ignore_index=True)
                    dataframe = dataframe.append(row_excel, ignore_index=True)
                    dataframe = dataframe.append(row_sql, ignore_index=True)
                iteration = 1
        except Exception as e:
            print(e)
            iteration = 1
            continue
    dataframe.to_excel("Data not properly transferred in SQL: DELETE, CHANGE Test3.xlsx")
    dataframe2.to_excel("Data not properly transferred SQL: CREATE in Test3.xlsx")
    print(dataframe, dataframe2)


value_checking(unmached_data, unmached_data2 )         




