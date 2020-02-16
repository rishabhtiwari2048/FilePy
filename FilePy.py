import cx_Oracle;
def filepy():
    print('Establishing database connection')
    ## Step 1 Connect to database
    username = "hr"
    password = "123"
    connection = cx_Oracle.connect(username,password,"LAPTOP-DO3LVILU/XEPDB1", encoding = "utf-8")
    print('Connection established successfully')
    ## Step 2 Execute the query to get the list of community ids
    ## Step 3 Get the data for each community and make a separate file of it
    ## Repeat the step
    ## Close the database connection
    connection.close()
    print("Connection closed successfully")
filepy()