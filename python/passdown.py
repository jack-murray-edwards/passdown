'''
Passdown_project__generator.py
V1.0
:author: Jack Murray
company: Edwards Vacuum. ses
jack.murray@edwardsvacuum.com
'''

import datetime
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import range_boundaries
from openpyxl.drawing.image import Image

# get date range for generating dictionary
start_day = datetime.date(2023, 4, 3)
end_day = datetime.date(2023, 7, 7)
# TODO Get this info from somewhere other than the code

def get_workdays(start_date, end_date):

    '''
    return a list of the dates of all the workdays between start_date and end_date

    :param start_date: a datetime.date object
    :param end_date: a datetime.date object
    :returns: a list of dictionaries in the form {'date': datetime.date, 'week_num': int}
    :rtype: list
    :requires: datetime
    :date: 2023-03-22
    '''

    workdays = []
    current_date = start_date

    while current_date <= end_date:
        if current_date.weekday() < 5:  # Monday = 0, Friday = 4
            workdays.append([current_date,
                                current_date.isocalendar()[1],
                                current_date.strftime("%m-%d")])

        current_date += datetime.timedelta(days=1)

    return workdays

def create_workbook(sheet_names, filename):
    '''
    Create a new excel workbook with the given sheet names

    :param sheet_names: a list of strings containing the names of the sheets to create
    :param filename: a string containing the name of the file to save the workbook to (including the file extension)
    :requires: openpyxl
    :date: 2023-03-22
    '''

    # Create a new workbook
    workbook = openpyxl.Workbook()

    # Rename the default sheet to the first sheet name in the list
    workbook.active.title = sheet_names[0]

    # Create the remaining sheets using the remaining sheet names in the list
    for name in sheet_names[1:]:
        workbook.create_sheet(title=name)

    # Save the workbook to a file
    workbook.save(filename)

def get_sheet_names(workdays):
    '''
    Create a list of strings containing the names of the sheets to create

    :param workdays: a list of dictionaries in the form {'date': datetime.date, 'week_num': int}
    :returns: a list of strings containing the names of the sheets to create
    :rtype: list
    :requires: openpyxl
    :date: 2023-03-22
    '''

    sheet_names = []
    sheet_names.append("Contents")
    sheet_names.append("Passdown Summary")
    for day in workdays:
        sheet_names.append(day[2])
    sheet_names.append("passdown_assets")

    return sheet_names

def copy_template_to_sheets(template_filename, sheet_range):
    '''
    Copy the template sheet to the specified sheets in the workbook
    
    :param template_filename: a string containing the name of the file to load the workbook from (including the file extension)
    :param sheet_range: a range of sheet names to copy the template to
    :requires: openpyxl
    :date: 2023-03-22
    '''

    # Load the template file
    wb = openpyxl.load_workbook(template_filename)

    # Get the template sheet
    template_sheet = wb.active

    # Loop over the sheets in the specified range
    for sheet_name in wb.sheetnames[sheet_range]:
        # Skip the template sheet
        if sheet_name == template_sheet.title:
            continue

        # Copy the template sheet to the current sheet
        target_sheet = wb[sheet_name]
        target_sheet.delete_rows(1, target_sheet.max_row)
        for row in template_sheet.iter_rows():
            for cell in row:
                target_sheet[cell.coordinate].value = cell.value
                target_sheet[cell.coordinate].number_format = cell.number_format
                target_sheet[cell.coordinate].font = cell.font
                target_sheet[cell.coordinate].border = cell.border
                target_sheet[cell.coordinate].fill = cell.fill
                target_sheet[cell.coordinate].alignment = cell.alignment

    # Save the modified workbook
    wb.save(template_filename)

def create_daily_sheets(filename, workdays):
    '''
    Create a new excel workbook with the given sheet names

    :param filename: a string containing the name of the file to load the workbook from (including the file extension)
    :param workdays: a list of dictionaries in the form {'date': datetime.date, 'week_num': int}
    :requires: openpyxl
    :date: 2023-03-22
    '''

    # Load the workbook
    wb = openpyxl.load_workbook(filename)

    # Loop over the workdays
    for day in workdays:
        # Get the sheet for the current day
        sheet = wb[day[2]]
        

        #resize the columns
        sheet.column_dimensions["A"].width = 11
        sheet.column_dimensions["B"].width = 17
        sheet.column_dimensions["C"].width = 43
        sheet.column_dimensions["D"].width = 65
        sheet.column_dimensions["E"].width = 8.75
        sheet.column_dimensions["F"].width = 10.5
        sheet.column_dimensions["G"].width = 22.5
        sheet.column_dimensions["H"].width = 22.75
        sheet.column_dimensions["I"].width = 22.75
        sheet.column_dimensions["J"].width = 3.5
        sheet.column_dimensions["K"].width = 11
        sheet.column_dimensions["L"].width = 8.75
        sheet.column_dimensions["M"].width = 14.5
        sheet.column_dimensions["N"].width = 71
        sheet.column_dimensions["O"].width = 12
        sheet.column_dimensions["P"].width = 34

        #create large boarder around the sheet
        # sheet["A1:P103"].border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style="thick"),
        #                                                 left=openpyxl.styles.Side(border_style="thick"),
        #                                                 right=openpyxl.styles.Side(border_style="thick"),
        #                                                 top=openpyxl.styles.Side(border_style="thick"))
        
        #create sheet title
        sheet.merge_cells("A1:P1")
        sheet.row_dimensions[1].height = 84
        sheet["A1"] = "Project Passdown for Work Week " + str(day[1]) + " - " + day[0].strftime("%A, %B %d, %Y")
        sheet["A1"].font = openpyxl.styles.Font(bold=True, size=20)
        sheet["A1"].alignment = openpyxl.styles.Alignment(horizontal="center", vertical="bottom")
        sheet["A1"].fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="C0C0C0")
        for row in sheet.iter_rows(min_row=1, max_row=1, min_col=1, max_col=16):
            for cell in row:
                cell.border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style="thick"))
        #Add the logo
        img = openpyxl.drawing.image.Image(
            '/home/jack/Dropbox/Workspace/code/passdown/python/Edwards_logo_for_sheets.png')
        sheet.add_image(img, "H1")

        # create edwards contact info
        sheet.merge_cells("A2:P2")
        sheet["A2"] = "Edwards Project Manager: Robert Nolan (503)753-0590   -   Edwards Site Manager: Joseph Baca (505)975-9464"
        sheet["A2"].font = openpyxl.styles.Font(size=14)
        sheet["A2"].alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
        sheet["A2"].fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="C0C0C0")
        for row in sheet.iter_rows(min_row=2, max_row=2, min_col=1, max_col=16):
            for cell in row:
                cell.border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style="thick"))

        #create passdown and look ahead headers
        sheet.merge_cells("A3:I3")
        sheet["A3"] = "Passdown"
        sheet["A3"].font = openpyxl.styles.Font(bold=True, size=16)
        sheet["A3"].alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
        sheet["A3"].border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style="thick"))
        sheet["A3"].fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="808080")
        
        sheet.merge_cells("K3:P3")
        sheet["K3"] = "Look Ahead"
        sheet["K3"].font = openpyxl.styles.Font(bold=True, size=16)
        sheet["K3"].alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
        sheet["K3"].border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style="thick"))
        sheet["K3"].fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="808080")
        
        #create the divider between passdown and look ahead
        sheet.merge_cells("J3:J103")
        sheet["J3"].fill = openpyxl.styles.PatternFill(
            patternType="solid", fgColor="000000")
        
        #header styles
        

        #create the passdown headers
        sheet["A4"] = "Task ID"
        sheet["B4"] = "Team Assigned"
        sheet["C4"] = "Task"
        sheet["D4"] = "Outcome"
        sheet["E4"] = "SMA"
        sheet["F4"] = "KAWIs"
        sheet["G4"] = "Next Step"
        sheet["H4"] = "AIO Request"
        sheet["I4"] = "AIO Completion"

        # passdown header formatting
        for row in sheet.iter_rows(min_row=4, max_row=4, min_col=1, max_col=9):
            for cell in row:
                cell.font = openpyxl.styles.Font(bold=True, size=11)
                cell.border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style="thick"))

        #create the look ahead headers
        sheet["K4"] = "Task ID"
        sheet["L4"] = "Priority Rank"
        sheet["M4"] = "Estimated Completion"
        sheet["N4"] = "Task"
        sheet["O4"] = "Time"

        # Look ahead headers formatting
        for row in sheet.iter_rows(min_row=4, max_row=4, min_col=11, max_col=15):
            for cell in row:
                cell.font = openpyxl.styles.Font(bold=True, size=11)
                cell.border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style="thick"))

        #protect the sheet and allow modifications to the passdown and look ahead sections only
        sheet.protection.sheet = True
        sheet.protection.set_password("!Edwards!")
        sheet.protection.enable()

        # enable editing for passdown
        for row in sheet.iter_rows(min_row=4, max_row=103, min_col=1, max_col=9):
            for cell in row:
                cell.protection = openpyxl.styles.Protection(locked=False)

        # enable editing for look ahead
        for row in sheet.iter_rows(min_row=4, max_row=103, min_col=11, max_col=15):
            for cell in row:
                cell.protection = openpyxl.styles.Protection(locked=False)

        #create passdown and look ahead table
        passdown_display_name = day[2] + "_Passdown"
        look_ahead_display_name = day[2] + "_Look_Ahead"
        passdown_table = Table(ref='A4:I103', displayName=passdown_display_name, headerRowCount=1)
        look_ahead_table = Table(ref='K4:O103', displayName=look_ahead_display_name, headerRowCount=1)


        sheet_table_style = TableStyleInfo(
            name="TableStyleMedium15", showRowStripes=True, showColumnStripes=False)
        
        passdown_table.tableStyleInfo = sheet_table_style
        look_ahead_table.tableStyleInfo = sheet_table_style

        sheet.add_table(passdown_table)
        sheet.add_table(look_ahead_table)

        #change header style
        header_style = openpyxl.styles.Font(bold=True, size=11, color="e5e4e2")
        for row in sheet.iter_rows(min_row=4, max_row=4, min_col=1, max_col=15):
            for cell in row:
                cell.font = header_style

        #create passdown column borders
        for col in sheet.iter_cols(min_row=4, max_row=103, min_col=1, max_col=9):
            for cell in col:
                cell.border = openpyxl.styles.Border(right=openpyxl.styles.Side(border_style="thick"))

        #create look ahead column borders
        for col in sheet.iter_cols(min_row=4, max_row=103, min_col=11, max_col=16):
            for cell in col:
                cell.border = openpyxl.styles.Border(right=openpyxl.styles.Side(border_style="thick"))

        #create bottom boarder
        for row in sheet.iter_rows(min_row=103, max_row=103, min_col=1, max_col=15):
            for cell in row:
                cell.border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(border_style="thick"))



    # Save the modified workbook
    wb.save(filename)

def create_contents_sheet(filename):
    '''
    Create a contents sheet in the workbook with links to all other sheets

    :param filename: a string containing the name of the file to load the workbook from (including the file extension)
    :requires: openpyxl
    :date: 2023-03-22
    '''
    # Load the workbook
    wb = openpyxl.load_workbook(filename)

    # Create a new sheet for the contents
    contents_sheet = wb.worksheets[0]

    # Loop over the sheets and add links to the contents sheet
    for sheet in wb.sheetnames:
        # Skip the contents sheet
        if sheet == "Contents":
            continue

        # Add a hyperlink to the sheet
        cell = contents_sheet.cell(row=contents_sheet.max_row+1, column=1)
        cell.value = sheet
        cell.hyperlink = f"'{sheet}'!A1"

    # Save the modified workbook
    wb.save(filename)


project_workbook = "excel/Project_Passdown_WW14-WW28.xlsx"
my_workdays = get_workdays(start_day, end_day)
sheet_names = get_sheet_names(my_workdays)
create_workbook(sheet_names, project_workbook)
create_contents_sheet(project_workbook)
create_daily_sheets(project_workbook, my_workdays)

