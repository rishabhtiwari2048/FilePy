import cx_Oracle
import xlsxwriter


def filepy():
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
    records = connection.cursor()
    for (community_id, community_name,) in communities.execute(
            "SELECT distinct department_id, department_name FROM t_demo"):
        workbook = xlsxwriter.Workbook("E:/A/" + community_name + ".xlsx",{'constant_memory':True})
        print('****************************************')
        print(community_name+".xls created successfully")
        worksheet = workbook.add_worksheet(community_name)
        row_num = 0
        col_num = 0
        print('-----------------------------------------')
        print("writing data")
        cell_format1 = workbook.add_format()
        cell_format1.set_bold(True)
        cell_format1.set_bg_color('cyan')
        cell_format1.set_font_size(12)
        cell_format1.set_font_color('black')
        cell_format1.set_border(1)
        cell_format1.set_font('Calibri')
        worksheet.write_row(row_num,col_num,file_header,cell_format1)
        row_num = row_num + 1
        query = "SELECT * FROM t_demo WHERE department_id =" + str(community_id)
        records.execute(query)
        rows = records.fetchall()
        cell_format2 = workbook.add_format()
        cell_format2.set_border(1)
        cell_format2.set_font_color('black')
        cell_format2.set_font('Calibri')
        cell_format2.set_font_size(10)
        for row in rows:
            worksheet.write_row(row_num,col_num,row,cell_format2)
            row_num = row_num + 1
        print(str(row_num)+" rows written")
        print("writing complete")
        print("------------------------------------------")
        print("Closing "+community_name+".xls")
        print('*******************************************')
        workbook.close()
    connection.close()
    print("Connection closed successfully")


filepy()
