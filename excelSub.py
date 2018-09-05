import win32com.client

print("here")

office = win32com.client.Dispatch("Excel.Application")
wb = office.Workbooks.Open(r"C:\Users\Bobli\PycharmProjects\ExcelService\sales.xlsx")
ws = wb.Worksheets
active_sheet = ws[0]
pivots = active_sheet.PivotTables()
active_pivot = pivots[0]
pivot_fields = active_pivot.PivotFields()
temp = str(pivot_fields[0])

wb.Close(True)
office.Quit()

with open(r"C:\Users\Bobli\PycharmProjects\ExcelService\test.txt", 'w') as file:
    file.write(temp)

print("there")