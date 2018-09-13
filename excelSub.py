import win32com.client
from datetime import datetime
import time
import json

# with open('task.json') as file:
#     request = json.load(file)
#
# if request["operation"] == "request_fields":
#     office = win32com.client.Dispatch("Excel.Application")
#     wb = office.Workbooks.Open(r"C:\Users\Bobli\PycharmProjects\ExcelService\sales.xlsx")
#     response = {
#         "operation": "request_fields",
#         "array": [
#             1,
#             2,
#             3
#         ],
#         "write_ts": "Hello World"
#     }
#     with open('task_response.json', 'w') as file:
#         file.write("Request acknowledge")
#     wb.Close(True)
#     office.Quit()

start = str(datetime.now())

office = win32com.client.Dispatch("Excel.Application")
wb = office.Workbooks.Open(r"C:\Users\Bobli\PycharmProjects\ExcelService\sales.xlsx")
ws = wb.Worksheets
active_sheet = ws[0]
pivots = active_sheet.PivotTables()
active_pivot = pivots[0]
pivot_fields = active_pivot.PivotFields()

# for field in pivot_fields:
#     print(field.Value)
temp = active_pivot.CubeFields[1].Value
print(temp)

# numOfFields = len(table.PivotFields())
# for i in range (0, numOfFields):
#     print(table.PivotFields()[i].Caption + " :::: " + table.PivotFields()[i].Value)
# return table.PivotFields()

wb.Close(True)
office.Quit()

end = str(datetime.now())


print("there")
