import xlsxwriter
import os
import Pull_RawData

def write_employee(worksheet, employee, weekly_report, row, col, data_format, name_format, needs_attention_format,
                   percent_format):

    smi_internal_total = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.SMI_INTERNAL]
    shine_sys_internal_total = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.SHINE_SYS_INTERNAL]
    my_vv_internal = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.MY_VV_INTERNAL]
    si_internal_total = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.SI_INTERNAL]
    shine_family_allocation = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.SHINE_FAMILY_ALLOCATION]
    pto_holiday_total = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.PTO_FLOATING_HOLIDAY]


    worksheet.write(row, col, employee.name, name_format)
    col += 1

    # print headers
    for i in range(col, 12):
        worksheet.write(row, col, '', name_format)
        col += 1
    row += 1
    # print projects
    for project in weekly_report.proj_lst:
        col = 1
        worksheet.write(row, col, project.name, data_format)
        col += 1
        worksheet.write(row, col, project.hours, data_format)

        col += 1
        # print percentages
        # print percentage totals
        for index in range(len(project.proj_hour_breakdown_lst) - 1 ):
        # for num in project.proj_hour_breakdown_lst:
            num = project.proj_hour_breakdown_lst[index]
            value_to_write = num / weekly_report.total_hours()
            worksheet.write(row, col, value_to_write, percent_format)
            col += 1

        col = 1
        row += 1

    # print subtotal
    col = 0
    worksheet.write(row, col, 'Subtotal', data_format)
    col = 2
    worksheet.write(row, col, weekly_report.total_hours(), needs_attention_format)
    col += 1
    # print percentage totals

    for num in weekly_report.weekly_hour_breakdown_lst:
        value_to_write = num / weekly_report.total_hours()
        worksheet.write(row, col, value_to_write, percent_format)
        col += 1
    # print total internal
    total_internal = smi_internal_total + shine_sys_internal_total + my_vv_internal + si_internal_total
    worksheet.write(row, col, )

    row += 1

    return row

def Create_Excel_File(report_file_path, employees, report_name_lst):


    workbook = xlsxwriter.Workbook(report_file_path)
    headers = ['User', 'Project', 'Hours', 'SMI Internal', 'SHINE SYS Internal', 'My VV Internal', 'SI Internal',
               'SHINE Family Allocation', 'PTO/Floating / Holiday', 'Total Internal', 'SHINE Companies', 'External',
                'Total']

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

        percent_format = workbook.add_format()
        percent_format.set_text_wrap()
        percent_format.set_num_format('0.00%')

        name_format = workbook.add_format()
        name_format.set_text_wrap()
        name_format.set_top()

        needs_attention_format = workbook.add_format()
        needs_attention_format.set_text_wrap()
        needs_attention_format.set_bg_color('yellow')


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
                    row = write_employee(worksheet,employee,weekly_report,row,col,normal_format, name_format,
                                         needs_attention_format, percent_format)


                    break


    workbook.close()
    os.startfile(report_file_path)