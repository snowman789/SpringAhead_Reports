import xlsxwriter
import os
import Pull_RawData

def write_employee(worksheet, employee, weekly_report, row, col, data_format, name_format):

    worksheet.write(row, col, employee.name, name_format)
    col += 1


    for i in range(col, 12):
        worksheet.write(row, col, '', name_format)
        col += 1
    row += 1
    for project in weekly_report.proj_lst:
        col = 1
        worksheet.write(row, col, project.name, data_format)
        col += 1
        worksheet.write(row, col, project.hours, data_format)


        col = 1
        row += 1

    return row

def Create_Excel_File(report_file_path, employees, report_name_lst):


    workbook = xlsxwriter.Workbook(report_file_path)
    headers = ['User', 'Project', 'Hours', 'SMI Internal', 'My VV Internal', 'SI Internal',
               'SHINE Family Allocation', 'PTO/FLoating/Holiday', 'Total Internal', 'SHINE Companies',
               'External', 'Total']

    for i in range(0,len(report_name_lst)):

        from_time = report_name_lst[i][0]
        to_time = report_name_lst[i][1]

        from_time_str = report_name_lst[i][0].strftime("%m-%d-%y")
        to_time_str = report_name_lst[i][1].strftime("%m-%d-%y")
        name = "Report " + str(from_time_str) + " to " + str(to_time_str)


        worksheet = workbook.add_worksheet(name)
        worksheet.freeze_panes(1,0)
        worksheet.set_column(2, 15,12)
        worksheet.set_column('A:A', 19)
        worksheet.set_column('B:B', 55)

        header_format = workbook.add_format()
        header_format.set_bold()
        header_format.set_bg_color('#add8e6')
        header_format.set_center_across()
        header_format.set_text_wrap()

        header_format.set_border()

        normal_format = workbook.add_format()
        normal_format.set_text_wrap()

        name_format = workbook.add_format()
        name_format.set_text_wrap()
        name_format.set_top()


        row=0
        col=0

        for item in headers:
            worksheet.write(row, col, item, header_format)
            col += 1
        col = 0
        row += 1

        #     iterate through employees
        hits = 0
        for employee in employees:
            # this is used in case some employees are added part way through the cycle
            weekly_index = employee.weekly_index
            for index in range(weekly_index, len(employee.weekly_reports_lst)):
                weekly_report = employee.weekly_reports_lst[index]
                if weekly_report.from_date == from_time and weekly_report.to_date == to_time:

                    hits += 1
                    employee.weekly_index = weekly_index + 1
                    row = write_employee(worksheet,employee,weekly_report,row,col,normal_format, name_format)


                    break


    workbook.close()
    os.startfile(report_file_path)