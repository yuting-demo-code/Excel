import xlwings as xw

def main():
    # Example: Accessing a worksheet and a cell
    wb = xw.Book.caller()
    sheet = wb.sheets['工作表1']

    # Example: Modifying a cell value
    sheet.range('B1').value = 'Hello, Excel!'

