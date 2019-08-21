import csv
import datetime
from enum import Enum
import os
import configparser
import Create_Excel_Sheet

employee_lst = []
report_name_lst = []

class Proj_category(Enum):
    UNKNOWN = 0
    SMI_INTERNAL = 1
    SHINE_SYS_INTERNAL = 2
    MY_VV_INTERNAL = 3
    SI_INTERNAL = 4
    SHINE_FAMILY_ALLOCATION = 5
    PTO_FLOATING_HOLIDAY = 6

class Weekly_Report:
    def __init__(self, from_date, to_date):
        self.from_date = from_date
        self.to_date = to_date
        self.proj_lst = []
        self.weekly_hour_breakdown_lst = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 ] #see enum to find indices

class Project:
    def __init__(self, name, hours):
        self.name = name
        self.code = 'no code defined'
        self.hours = float(hours)
        self.proj_category = Proj_category.UNKNOWN

class Employee:
    def __init__(self, name):
        self.name = name

        self.projects = []
        self.total_hours = 0
        self.weekly_reports_lst = []
        self.hour_breakdown_lst = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 ] #see enum to find indices
        self.weekly_index = 0

    def add_project(self, project):
        project_exists = False

        existing_project = project
        for obj in self.projects:
            if obj.name == project.name:
                project_exists = True
                existing_project = obj


        if project_exists:
            existing_project.hours += project.hours

        else:

            self.projects.append(project)


def assign_project_category(config, name):
    temp_lst = name.split()

    code = temp_lst[0]


    #consider moving this outside function if program speed is an issue

    smi_internal = str(config['SMI_INTERNAL']).split(',')
    shine_sys = str(config['SHINE_SYS_INTERNAL']).split(',')
    my_vv = str(config['MY_VV_INTERNAL']).split(',')
    si = str(config['SI_INTERNAL']).split(',')
    shine_family_allocation = str(config['SHINE_FAMILY_ALLOCATION']).split(',')
    pto_floating_holiday = str(config['PTO_FLOATING_HOLIDAY']).split(',')

    if str_in_lst(code, smi_internal):
        return Proj_category.SMI_INTERNAL
    elif str_in_lst(code, shine_sys):
        return Proj_category.SHINE_SYS_INTERNAL
    elif str_in_lst(code, my_vv):
        return Proj_category.MY_VV_INTERNAL
    elif str_in_lst(code, si):
        return Proj_category.SI_INTERNAL
    elif str_in_lst(code, shine_family_allocation):
        return Proj_category.SHINE_FAMILY_ALLOCATION
    elif str_in_lst(code, pto_floating_holiday):
        return Proj_category.PTO_FLOATING_HOLIDAY
    else:
        return Proj_category.UNKNOWN

def str_in_lst(target, lst):
    target = str(target).lower()
    for word in lst:
        word = str(word).lower()
        word = word.strip()
        if word == target:
            return True
    return False



def get_Report_Date(row):
    date_row = str(row[0]).split()
    fromDate_str = str(date_row[0]).split('/')
    to_Date_str = str(date_row[3]).split('/')

    report_from_date = datetime.datetime(int(fromDate_str[2]), int(fromDate_str[0]), int(fromDate_str[1]))
    report_to_date = datetime.datetime(int(to_Date_str[2]), int(to_Date_str[0]), int(to_Date_str[1]))

    return[report_from_date, report_to_date]

def cell_Is_A_Name(cell):
    cell_str = str(cell)
    decoys = ['','Subtotal', 'Total']
    name_flag = True

    for word in decoys:
        if cell_str == str(word):
            name_flag = False
    return name_flag

def employee_Exists(name):
    employee_Exists = False
    for employee in employee_lst:
        if employee.name == name:
            employee_Exists = True
    return employee_Exists

def get_employee(name):
    for employee in employee_lst:
        if employee.name == name:
            return employee

def open_files(path):
    counter = 0
    for filename in os.listdir(path):
        if filename.endswith(".csv"):
            dir_path = (os.path.join(path, filename))
            pull_data(dir_path)
            counter += 1
    return counter


def pull_data(path):

    #get user config
    configFilePath = os.getcwd() + '\config.txt'
    configParser = configparser.RawConfigParser()
    configParser.read(configFilePath)
    config = configParser['DEFAULT']



    raw_data = open(str(path), encoding = 'utf8')
    csv_raw_data = csv.reader(raw_data)

    counter = 1
    employee = Employee("temp")

    report_name_saved = False


    for row in csv_raw_data:


        #pull date for report
        if counter == 5:
            dates = get_Report_Date(row)
            report_From_Date = dates[0]
            report_To_Date = dates[1]
            if not report_name_saved:
                report_name_lst.append([report_From_Date,report_To_Date])


        # where data starts
        if counter >= 9:
            #name_found
            if(cell_Is_A_Name(row[0])):
                name = row[0]

                #create new employee and add to employee list
                if not employee_Exists(name):
                    employee = Employee(name)
                    employee.weekly_reports_lst.append(Weekly_Report(report_From_Date,report_To_Date))
                    employee_lst.append(employee)
                else:
                    employee = get_employee(name)
                    employee.weekly_reports_lst.append(Weekly_Report(report_From_Date, report_To_Date))

            # This elif ensures there's actually a project in the row before trying to manipulate the data
            elif(row[0] != 'Total' and row[0] != 'Subtotal'):

                # create project object for this row's data and append it to weekly report for this employee
                project_name = row[1]
                hours = float(row[2])
                project = Project(project_name,hours)
                project.proj_category = assign_project_category(config, project_name)
                weekly_report = employee.weekly_reports_lst[-1]

                if project.proj_category == Proj_category.SHINE_FAMILY_ALLOCATION:

                    weekly_report.weekly_hour_breakdown_lst[Proj_category.SMI_INTERNAL.value] += (hours * .25)
                    weekly_report.weekly_hour_breakdown_lst[Proj_category.SHINE_SYS_INTERNAL.value] += (hours * .74)
                    weekly_report.weekly_hour_breakdown_lst[Proj_category.SI_INTERNAL.value] += (hours * .01)
                else:
                    weekly_report.weekly_hour_breakdown_lst[project.proj_category.value] += hours

                # append project
                weekly_report.proj_lst.append(project)

                # record data for summary report
                if project.proj_category == Proj_category.SHINE_FAMILY_ALLOCATION:
                    employee.hour_breakdown_lst[Proj_category.SMI_INTERNAL.value] += (hours * .25)
                    employee.hour_breakdown_lst[Proj_category.SHINE_SYS_INTERNAL.value] += (hours * .74)
                    employee.hour_breakdown_lst[Proj_category.SI_INTERNAL.value] += (hours * .01)
                else:
                    employee.hour_breakdown_lst[project.proj_category.value] += hours
                employee.total_hours += hours
                employee.add_project(project)












        counter +=1



    return employee_lst

# ----------------------------------------------------------------------
if __name__ == "__main__":
    path = r'C:\Users\isaac\OneDrive\Desktop\Shine_Systems\Raw Data'
    time = datetime.datetime.now().strftime("%m-%d-%y")

    time = r'\Test report from '+ time

    report_path = r'C:\Users\isaac\OneDrive\Desktop\Shine_Systems\SpringAhead_test_reports' + time +'.xlsx'
    num_reports = open_files(path)

    employee = employee_lst[1]

    Create_Excel_Sheet.Create_Excel_File(report_path, employee_lst, report_name_lst)
    weekly_report = employee.weekly_reports_lst[0]
    project = weekly_report.proj_lst[0]
    print("Employee name: " + str(employee.name))
    print("Employee's total hours: " + str(employee.total_hours))
    print("Project Name : " + str(project.name))
    print("project category: " + str(project.proj_category))
    print("Num hours: " + str(project.hours))

    for project in employee.projects:
        print(project.name + "    hours: " + str(project.hours))

