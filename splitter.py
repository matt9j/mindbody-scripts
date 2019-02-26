from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

PAYROLL_PERIOD="Feb 2019"
# Load the excel sheet file into python
workbook = load_workbook('./test.xlsx')

# Access the first sheet in the xlsx file.
# Note that in the future if mindbody changes their format, this may
# also need to change.
source_sheet = workbook[workbook.sheetnames[0]]

# Returns true if the row is blank
def row_is_blank(row_index):
    start = "A" + row_index
    end = get_column_letter(source_sheet.max_column) + row_index
    print(start)
    print(end)
    print("starting the loop")
    for cellobj in source_sheet[start:end]:
        for cell in cellobj:
            print (cell.value)
            if cell.value is not None:
                return False

    return True

# Writing test
import xlwt

sheet1 = book.add_sheet("sheet1")

# Copies a row from the source sheet into the new sheet
# Rows are specified in one indexed form.
def copy_row_to_sheet(new_sheet, source_row_number, dest_row_number):
    start = "A" + str(source_row_number)
    end = get_column_letter(source_sheet.max_column) + str(source_row_number)
    for cellobj in source_sheet[start:end]:
        for cell in cellobj:
            # The write library is zero indexed, but the read library
            # is one indexed... facepalm... settle one one indexed for
            # now.
            new_sheet.write(dest_row_number - 1, cell.column - 1, cell.value)


copy_row_to_sheet(sheet1, 2, 1)

#loop over all rows in the source sheet.
#find blank rows, after each blank row...
#get the name, which will be the next row, first cell
#create a new book with the new name
#copy rows into the new book
#when find blank row, save the book, start over.
book = None
current_name = None
for source_row_number in range(1, source_sheet.max_row + 1):
    if row_is_blank(source_row_number):
        if book is not None:
            book.save(current_name + " - " + PAYROLL_PERIOD)
        book = xlwt.Workbook(encoding="utf-8")
        current_name = 
        pass
    else:
        copy_row_to_sheet(new_sheet, source_row_number, dest_index)
        dest_index += 1
