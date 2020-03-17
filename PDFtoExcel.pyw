import os
import re
import sys
import ctypes
from openpyxl import Workbook


def messagebox(title,message):
    MessageBox = ctypes.windll.user32.MessageBoxW
    MessageBox(None, message, title, 0)


def pdftotext(filepath):
    if not filepath.endswith('.pdf'):
        messagebox('Not a pdf file','Please input a file in pdf format.')
        exit()
    if not os.path.isfile("pdftextExtracter.exe "):
        messagebox('PDF extractor not found', 'Please ensure pdftextextracter.exe is in the same folder as this script.')
        exit()
    cmd1 = f"pdftextExtracter.exe -layout {filepath}"
    os.system(cmd1)
    filepathtxt = filepath.replace(".pdf", ".txt")
    file = open(filepathtxt, 'r')
    lines = file.readlines()

    nameid = []
    nameidregex = re.compile(r'(\d+)\s+(.+) -.+')
    output = [["Employee Number", "Employee Name", "Designation", "Regular Hours", "Overtime Hours", "Regular Pay", "Overtime Pay"]]
    job_totalsregex_5 = re.compile(r'(\d+\s-\s.+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+.+\.\d+')
    job_totalsregex_4 = re.compile(r'(\d+\s-\s.+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+.+\.\d+')
    for row_num, line in enumerate(lines):
        if "Employee # And Name" in line or "Total Hours Worked This Pay Period:" in line:
            employeeIDandName = re.search(nameidregex, lines[row_num + 3])
            if employeeIDandName:
                employeeID = employeeIDandName[1]
                employeeName = employeeIDandName[2].strip()
                nameid.append([employeeID, employeeName])
        if "Job Totals" in line:
            iter_num = row_num
            reg_hours = 0
            overtime_hours = 0
            regular_pay = 0
            overtime_pay = 0
            while "Total Hours Worked This Pay Period:" not in lines[iter_num] and iter_num-row_num < 14:
                totals = re.search(job_totalsregex_5, lines[iter_num + 1])
                if totals:
                    designation = totals[1]
                    reg_hours = totals[2]
                    overtime_hours = totals[3]
                    regular_pay = totals[4]
                    overtime_pay = totals[5]
                    output.append([employeeID, employeeName, designation, reg_hours, overtime_hours, regular_pay, overtime_pay])
                else:
                    totals_2 = re.search(job_totalsregex_4, lines[iter_num + 1])
                    if totals_2:
                        designation = totals_2[1]
                        reg_hours = totals_2[2]
                        overtime_hours = "0"
                        regular_pay = totals_2[3]
                        overtime_pay = totals_2[4]
                        output.append([employeeID, employeeName, designation, reg_hours, overtime_hours, regular_pay, overtime_pay])
                iter_num += 1
    if not output:
        messagebox('Nothing to output', 'File was processed, but no data was exported.')
    filepathexcel = filepathtxt.replace(".txt", ".xlsx")
    if os.path.isfile(filepathexcel):
        messagebox2 = ctypes.windll.user32.MessageBoxW
        return_value = messagebox2(None, "The output excel file of the given PDF already exists. Click yes to rewrite, No to Cancel.", "File already exists", 4)
        '''
        MB_OK = 0
        MB_OKCANCEL = 1
        MB_YESNOCANCEL = 3
        MB_YESNO = 4

        IDOK = 0
        IDCANCEL = 2
        IDABORT = 3
        IDYES = 6
        IDNO = 7
        '''
        if return_value == 6:
            save_to_excel(output,filepathexcel)
        elif return_value == 7:
            exit()
    else:
        save_to_excel(output,filepathexcel)


def save_to_excel(output,filename):
    if output:
        excelfile = Workbook()
        excelsheet = excelfile.active
        excelsheet.title = "Output"
        for row in output:
            excelsheet.append(row)
        excelfile.save(filename)

if __name__ == "__main__":
    if not len(sys.argv) == 2:
        messagebox('PDF file not found', 'Please drag the pdf file to this script.')
        exit()
    filepath = sys.argv[1]
    pdftotext(filepath)