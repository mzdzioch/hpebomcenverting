from openpyxl.styles import Alignment, Font
from expy import open_xls_as_xlsx as open_xls_as_xlsx

import re

def swap(sheets, columnfirst, columnsecond):
    for a in sheets[columnfirst]:
        new_cell = a.offset(column=6)
        new_cell.value = a.value

    for b in sheets[columnsecond]:
        new_cell_2 = b.offset(column=-1)
        new_cell_2.value = b.value

    for c in sheets[columnsecond]:
        c.value = c.offset(column=5).value
        c.offset(column=5).value = None

def cellalign(worksheet, column):
    for c in worksheet[column]:
        c.alignment = Alignment(horizontal='center')
        c.font = Font(bold=True)

def setcloumnwidth(worksheet):
    worksheet.column_dimensions['A'].width = 53
    worksheet.column_dimensions['B'].width = 60.6423
    worksheet.column_dimensions['C'].width = 20
    worksheet.column_dimensions['D'].width = 20
    worksheet.column_dimensions['E'].width = 30
    worksheet.column_dimensions['F'].width = 30

def setcloumnwidthforfirstsheet(worksheet):
    worksheet.column_dimensions['A'].width = 53
    worksheet.column_dimensions['B'].width = 40
    worksheet.column_dimensions['C'].width = 30
    worksheet.column_dimensions['D'].width = 30
    worksheet.column_dimensions['E'].width = 30
    worksheet.column_dimensions['F'].width = 30

def setwrapping(worksheet, column):
    for c in worksheet[column]:
        c.alignment = Alignment(wrap_text=True)

def removestringafterchar(worksheet, column, char):
    for c in worksheet[column]:
        if(c.value.find(char)):
            text = c.value
            c.value = text.split(char, 1)[0]

def findtext(worksheet, column):
    count = 1
    for c in worksheet[column]:
        if(re.match("HPE [\d, 0-9]{1,3}[GB|TB]", c.value)):
            if (re.search("Memory", c.value)):
                splitted = c.value.split(")", 1)
                c.value = '="' + splitted[0] + ') (in total "&D'+str(count) + '*' + str(findvolume(c.value)) + '&"' + findunit(c.value) + ')' + splitted[1] + '"'
            else:
                splitted = c.value.split(" ", 2)
                c.value = '="' + splitted[0] + splitted[1] + ' (in total "&D'+str(count) + '*' + str(findvolume(c.value)) + '&"' + findunit(c.value) + ') ' + splitted[2] + '"'
        else:
            print()
        count = count + 1

def findvolume(text):
    array = text.split(" ")
    for word in array:
        if(re.match("[\d, 0-9]{1,3}[GB|TB]", word)):
            return int(re.match(r'\d+', word).group())
    return 0

def findunit(text):
    array = text.split(" ")
    for word in array:
        if(re.match("[\d, 0-9]{1,3}TB", word)):
            return "TB"
        if(re.match("[\d, 0-9]{1,3}GB", word)):
            return "GB"
    return "null"

def run(filename, true=None):
    # wb = open_xls_as_xlsx('/home/michal/Downloads/DL380_ESX_Host.XLS')
    if(re.search(".XLS", filename)):
        wb = open_xls_as_xlsx(filename)
        source = wb.active
        sheets = wb.copy_worksheet(source)
        if(re.search("Configuration Summary", sheets['A1'].value)):
            #sheets = wb.get_sheet_by_name(wb.sheetnames[0])
            cellalign(sheets, 'D')
            swap(sheets, 'C', 'D')
            setcloumnwidth(sheets)
            setcloumnwidthforfirstsheet(source)
            removestringafterchar(sheets, 'C', '#')
            findtext(sheets, 'B')
        else:
            return False

        wb.save(filename)
        return True
    else:
        return False



