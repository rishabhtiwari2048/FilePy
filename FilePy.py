import cx_Oracle
import xlsxwriter
def createfile(community_id, community_name, location, connection, table_name):
    records = connection.cursor()
    workbook = xlsxwriter.Workbook(location + community_name + ".xlsx", {'constant_memory': True})
    print('****************************************')
    print(community_name + ".xls created successfully")
    worksheet = workbook.add_worksheet(community_name)
    row_num = 0
    col_num = 0
    file_header = (
        "department_id", "department_name", "manager_id", "location_id", "employee_id", "first_name", "last_name",
        "email",
        "hire_date", "salary")
    print('-----------------------------------------')
    print("writing data")
    cell_format1 = workbook.add_format()
    cell_format1.set_bold(True)
    cell_format1.set_bg_color('cyan')
    cell_format1.set_font_size(12)
    cell_format1.set_font_color('black')
    cell_format1.set_border(1)
    cell_format1.set_font('Calibri')
    worksheet.write_row(row_num, col_num, file_header, cell_format1)
    row_num = row_num + 1
    query = "SELECT * FROM " + table_name + " WHERE department_id =" + str(community_id)
    records.execute(query)
    rows = records.fetchall()
    cell_format2 = workbook.add_format()
    cell_format2.set_border(1)
    cell_format2.set_font_color('black')
    cell_format2.set_font('Calibri')
    cell_format2.set_font_size(10)
    for row in rows:
        worksheet.write_row(row_num, col_num, row, cell_format2)
        row_num = row_num + 1
    print(str(row_num) + " rows written")
    print("writing complete")
    print("------------------------------------------")
    print("Closing " + community_name + ".xls")
    print('*******************************************')
    workbook.close()

def createmultiplefile(community_id, community_name, location, connection, table_name, no_of_files):
    records = connection.cursor()
    print(no_of_files)
    for i in range(1,no_of_files+1):
        print("**"+str(i))
        workbook = xlsxwriter.Workbook(location+community_name+"_"+str(i)+ ".xlsx", {'constant_memory': True})
        print('****************************************')
        print(community_name+str(i)+ ".xls created successfully")
        worksheet = workbook.add_worksheet(community_name+str(i))
        row_num = 0
        col_num = 0
        file_header = (
            "department_id", "department_name", "manager_id", "location_id", "employee_id", "first_name", "last_name",
            "email",
            "hire_date", "salary")
        print('-----------------------------------------')
        print("writing data")
        cell_format1 = workbook.add_format()
        cell_format1.set_bold(True)
        cell_format1.set_bg_color('cyan')
        cell_format1.set_font_size(12)
        cell_format1.set_font_color('black')
        cell_format1.set_border(1)
        cell_format1.set_font('Calibri')
        worksheet.write_row(row_num, col_num, file_header, cell_format1)
        row_num = row_num + 1
        if i == 1:
            query = "SELECT * FROM " + table_name + " WHERE department_id =" + str(community_id) +" ORDER BY 1 FETCH FIRST 60000 ROWS ONLY"
        else:
            query = "SELECT * FROM " + table_name + " WHERE department_id =" + str(
                community_id) + " ORDER BY 1 OFFSET "+str((i-1)*60000)+" ROWS FETCH NEXT 60000 ROWS ONLY"
        print(query)
        records.execute(query)
        rows = records.fetchall()
        cell_format2 = workbook.add_format()
        cell_format2.set_border(1)
        cell_format2.set_font_color('black')
        cell_format2.set_font('Calibri')
        cell_format2.set_font_size(10)
        for row in rows:
            worksheet.write_row(row_num, col_num, row, cell_format2)
            row_num = row_num + 1
        print(str(row_num) + " rows written")
        print("writing complete")
        print("------------------------------------------")
        print("Closing " + community_name+str(i)+ ".xls")
        print('*******************************************')
        workbook.close()


def filepy(table_name,location):
    print('Establishing database connection')
    username = "hr"
    password = "123"
    file_header = (
        "department_id", "department_name", "manager_id", "location_id", "employee_id", "first_name", "last_name",
        "email",
        "hire_date", "salary")
    connection = cx_Oracle.connect(username, password, "LAPTOP-DO3LVILU/XEPDB1", encoding="utf-8")
    print('Connection established successfully')
    communities = connection.cursor()
    for (community_id, community_name,file_count) in communities.execute(
            "SELECT distinct department_id, department_name,count(*) FROM "+table_name+" GROUP BY department_id, department_name"):
        if file_count <= 60000 :
            createfile(community_id,community_name,location,connection, table_name)
        else:
            no_of_files = file_count/60000
            print(file_count)
            print(no_of_files)
            if no_of_files > no_of_files//60000:
                no_of_files = int(no_of_files)+ 1
                print(no_of_files)
                createmultiplefile(community_id,community_name,location,connection, table_name,no_of_files)
            else:
                if no_of_files == no_of_files//2:
                    no_of_files = int(no_of_files)
                    createmultiplefile(community_id, community_name, location, connection, table_name, no_of_files)

    connection.close()
    print("Connection closed successfully")


filepy("t_demo","E:/A/")
