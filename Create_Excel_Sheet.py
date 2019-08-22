import xlsxwriter
import os
import Pull_RawData

def write_employee(worksheet, employee, weekly_report, row, col, data_format, name_format, needs_attention_format,
                   percent_format,tan_name_format,tan_format):



    # Find weekly percentages
    smi_internal_total = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.SMI_INTERNAL.value]
    shine_sys_internal_total = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.SHINE_SYS_INTERNAL.value]
    my_vv_internal = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.MY_VV_INTERNAL.value]
    si_internal_total = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.SI_INTERNAL.value]
    shine_family_allocation = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.SHINE_FAMILY_ALLOCATION.value]
    pto_holiday_total = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.PTO_FLOATING_HOLIDAY.value]
    external_hours = weekly_report.weekly_hour_breakdown_lst[Pull_RawData.Proj_category.UNKNOWN.value]

    total_internal = (smi_internal_total + pto_holiday_total)/weekly_report.total_hours()
    total_shine_companies = (shine_sys_internal_total + my_vv_internal + si_internal_total)/weekly_report.total_hours()
    total_external = (external_hours)/weekly_report.total_hours()
    weekly_total_percent = total_internal + total_shine_companies + total_external

    worksheet.write(row, col, employee.name, tan_name_format)
    col += 1

    # print headers
    for i in range(col, 13):
        worksheet.write(row, col, '', name_format)
        col += 1
    row += 1
    # print projects
    for project in weekly_report.proj_lst:
        project.calculate_percentages(weekly_report.total_hours())
        col = 0
        worksheet.write(row,col,'',tan_format)
        col += 1
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
            if value_to_write > 0:
                worksheet.write(row, col, value_to_write, percent_format)
            col += 1

        # total percentages
        worksheet.write(row,col,project.internal_hours_percent, percent_format)
        col += 1

        worksheet.write(row,col, project.shine_companies_hours_percent, percent_format)
        col += 1

        worksheet.write(row, col, project.external_percent, percent_format)
        col += 1

        worksheet.write(row, col, project.total_percent, percent_format)


        col = 1
        row += 1

    # print subtotal
    col = 0
    worksheet.write(row, col, 'Subtotal', needs_attention_format)
    col = 2
    worksheet.write(row, col, weekly_report.total_hours(), needs_attention_format)
    col += 1
    # print percentage totals


    for index in range(0,len(weekly_report.weekly_hour_breakdown_lst)-1):
        num = weekly_report.weekly_hour_breakdown_lst[index]
        value_to_write = num / weekly_report.total_hours()

        worksheet.write(row, col, value_to_write, percent_format)
        col += 1
    # print total internal

    worksheet.write(row, col, total_internal, percent_format )
    col += 1

    worksheet.write(row, col, total_shine_companies, percent_format)
    col += 1

    worksheet.write(row, col, total_external, percent_format)
    col += 1

    worksheet.write(row, col, weekly_total_percent, percent_format)
    col+=1

    row += 1

    return row


def create_summary(worksheet, headers, header_format, first_report_date, last_report_date, employees, data_format, name_format, needs_attention_format,
                   percent_format, total_percent_format,tan_name_format,tan_format):

    row = 0
    col = 0
    # write worksheet headers
    for item in headers:
        worksheet.write(row, col, item, header_format)
        col += 1
    col = 0
    row += 1


    total_hours = 0
    for employee in employees:
        total_hours += employee.total_hours

        smi_internal_total = employee.hour_breakdown_lst[Pull_RawData.Proj_category.SMI_INTERNAL.value]
        shine_sys_internal_total = employee.hour_breakdown_lst[Pull_RawData.Proj_category.SHINE_SYS_INTERNAL.value]
        my_vv_internal = employee.hour_breakdown_lst[Pull_RawData.Proj_category.MY_VV_INTERNAL.value]
        si_internal_total = employee.hour_breakdown_lst[Pull_RawData.Proj_category.SI_INTERNAL.value]
        shine_family_allocation = employee.hour_breakdown_lst[
            Pull_RawData.Proj_category.SHINE_FAMILY_ALLOCATION.value]
        pto_holiday_total = employee.hour_breakdown_lst[Pull_RawData.Proj_category.PTO_FLOATING_HOLIDAY.value]
        external_hours = employee.hour_breakdown_lst[Pull_RawData.Proj_category.UNKNOWN.value]

        total_internal = (smi_internal_total + pto_holiday_total) / employee.total_hours
        total_shine_companies = (
                                        shine_sys_internal_total + my_vv_internal + si_internal_total) / employee.total_hours
        total_external = (external_hours) / employee.total_hours
        weekly_total_percent = total_internal + total_shine_companies + total_external

        # write employee headers
        worksheet.write(row, col, employee.name, tan_name_format)
        col += 1

        # print headers
        for i in range(col, 13):
            worksheet.write(row, col, '', name_format)
            col += 1
        row += 1

        for project in employee.projects:

            project.calculate_percentages(employee.total_hours)
            col = 0
            worksheet.write(row,col,'',tan_format)
            col += 1
            worksheet.write(row, col, project.name, data_format)
            col += 1
            worksheet.write(row, col, project.hours, data_format)

            col += 1

            # print percentages
            # print percentage totals

            for index in range(len(project.proj_hour_breakdown_lst) - 1):
                # for num in project.proj_hour_breakdown_lst:
                num = project.proj_hour_breakdown_lst[index]
                # print("num " + str(num) + " total hours " + str(employee.total_hours))
                value_to_write = num / employee.total_hours
                if value_to_write > 0:
                    worksheet.write(row, col, value_to_write, percent_format)
                col += 1

            # total percentages for each project
            worksheet.write(row, col, project.internal_hours_percent, percent_format)
            col += 1

            worksheet.write(row, col, project.shine_companies_hours_percent, percent_format)
            col += 1

            worksheet.write(row, col, project.external_percent, percent_format)
            col += 1

            worksheet.write(row, col, project.total_percent, percent_format)

            col = 0
            row += 1

        col = 0
        worksheet.write(row, col, 'Subtotal', needs_attention_format)
        col = 2
        worksheet.write(row, col, employee.total_hours, needs_attention_format)
        col += 1
        # print percentage totals

        for index in range(0, len(employee.hour_breakdown_lst) - 1):
            num = employee.hour_breakdown_lst[index]
            value_to_write = num / employee.total_hours
            worksheet.write(row, col, value_to_write, percent_format)
            col += 1


        worksheet.write(row, col, total_internal, total_percent_format)
        col += 1

        worksheet.write(row, col, total_shine_companies, total_percent_format)
        col += 1

        worksheet.write(row, col, total_external, total_percent_format)
        col += 1

        worksheet.write(row, col, weekly_total_percent, total_percent_format)
        col += 1

        row += 1
        col = 0
    # write total hours for summary
    # print headers
    for i in range(col, 13):
        worksheet.write(row, col, '', name_format)
        col += 1
    row += 1

    col = 1
    
    worksheet.write(row, col, 'Total Hours:', data_format)
    col += 1
    worksheet.write(row,col, total_hours, needs_attention_format)


    return 1

def Create_Excel_File(report_file_path, employees, report_name_lst):


    workbook = xlsxwriter.Workbook(report_file_path)
    headers = ['User', 'Project', 'Hours', 'SMI Internal', 'SHINE SYS Internal', 'My VV Internal', 'SI Internal',
               'SHINE Family Allocation', 'PTO/Floating / Holiday', 'Total Internal', 'SHINE Companies', 'External',
                'Total']



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

    total_percent_format = workbook.add_format()
    total_percent_format.set_text_wrap()
    total_percent_format.set_num_format('0.00%')
    total_percent_format.set_bg_color('yellow')

    tan_color = '#FFFFFF'
    tan_name_format = workbook.add_format( )
    # tan_name_format.set_text_wrap()
    tan_name_format.set_top()
    # tan_name_format.set_bg_color(tan_color)


    tan_format = workbook.add_format()
    tan_format.set_text_wrap()
    # tan_format.set_bg_color(tan_color)

    name_format = workbook.add_format()
    name_format.set_text_wrap()
    name_format.set_top()

    needs_attention_format = workbook.add_format()
    needs_attention_format.set_text_wrap()
    needs_attention_format.set_bg_color('yellow')

    first_report_date = 0
    last_report_date = 0

    for i in range(0,len(report_name_lst)):



        weekly_total_hours = 0
        from_time = report_name_lst[i][0]
        to_time = report_name_lst[i][1]

        from_time_str = report_name_lst[i][0].strftime("%m-%d-%y")

        to_time_str = report_name_lst[i][1].strftime("%m-%d-%y")
        name = "Report " + str(from_time_str) + " to " + str(to_time_str)
        # get date range for summary
        if i == 0:
            first_report_date = from_time_str
        last_report_date = to_time_str
        worksheet = workbook.add_worksheet(name)
        worksheet.freeze_panes(1, 0)
        worksheet.set_column(2, 15, 12)
        worksheet.set_column('A:A', 19)
        worksheet.set_column('B:B', 55)


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
                                         needs_attention_format, percent_format,tan_name_format,tan_format)

                    weekly_total_hours += weekly_report.total_hours()
                    break

        worksheet.write(row, col, 'Total', name_format)
        col+=1
        for i in range(col, 13):
            worksheet.write(row, col, '', name_format)
            col += 1
        worksheet.write(row, 2, weekly_total_hours, name_format)
        row += 1

    # CREATE SUMMARY
    name = 'Summary ' + first_report_date + ' to ' + last_report_date
    worksheet = workbook.add_worksheet(name)
    # worksheet.freeze_panes(1,0)
    worksheet.freeze_panes(1,0)
    worksheet.set_column(2, 15,12)
    worksheet.set_column('A:A', 19)
    worksheet.set_column('B:B', 55)
    create_summary(worksheet, headers, header_format, first_report_date, last_report_date, employees, normal_format,
                   name_format, needs_attention_format,
                   percent_format,total_percent_format,tan_name_format,tan_format)



    workbook.close()
    os.startfile(report_file_path)



