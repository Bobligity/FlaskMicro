import win32com.client
import datetime


class ExcelCom():

    def __init__(self):
        self.wb = self.openExistingWorkbook()

    def __del__(self):
        self.wb.Close(True)
        self.office.Quit()
        del self.office  # this line removed it from task manager in my case

    def openExistingWorkbook(self):
        self.office = win32com.client.Dispatch("Excel.Application")
        #self.setVisible(True)
        self.wb = self.office.Workbooks.Open(r"C:\Users\XDC8549\git\PythonExcelCom\RegressionDevBlank.xlsx")
        return self.wb

    def setVisible(self, flag):
        self.office.Visible = flag

    def getWorksheetCount(self):
        count = str(self.wb.Sheets.Count)
        self.writeLog(count + " sheet(s) found")
        return self.wb.Sheets.Count

    def getWorksheets(self, wb):
        self.worksheets = wb.Worksheets
        return self.worksheets

    def getWorksheet(self, sheet):
        self.ws= self.wb.Worksheets[sheet]
        return self.ws

    def getPivotTableCount(self, worksheet):
        ptc = worksheet.PivotTables().Count
        ptcprint = str(ptc) + " pivot(s) found"
        self.writeLog(ptcprint + " sheet(s) found")
        return ptc

    def getPivotTables(self, worksheet):
        return worksheet.PivotTables()

    def getPivotTable(self, index, worksheet):
        return worksheet.PivotTables()[index]

    def getActivePivotTableFields(self, table):
        numOfFields = len(table.PivotFields())
        for i in range (0, numOfFields):
            print(table.PivotFields()[i].Caption + " :::: " + table.PivotFields()[i].Value)
        return table.PivotFields()

    def getAllPivotTableFields(self,table):
        cubeFields = table.CubeFields
        numOfFields = len(cubeFields)
        #for i in range (0, numOfFields):
            #print cubeFields[i].Caption + " : " + str(i) + " : " + str(cubeFields[1].DragToHide)

        #print table.CubeFields[10].Caption
        return cubeFields

    pivotField = {'column':2,'row':1,'filter':3,'values':4,'hidden':0}

    def removeField(self, cubeFields, fieldName):
        numOfFields = len(cubeFields)
        for i in range (0, numOfFields):
            if cubeFields[i].Caption == fieldName:
                #print cubeFields[i].Caption + " : " + str(i) + " : " + str(cubeFields[1].DragToHide)
                cubeFields[i].Orientation = self.pivotField['hidden']
        return self.pivotField['hidden']

    def addFieldToRow(self, cubeFields, fieldName):
        numOfFields = len(cubeFields)
        for i in range(0, numOfFields):
            if cubeFields[i].Caption == fieldName:
                # print cubeFields[i].Caption + " : " + str(i) + " : " + str(cubeFields[1].DragToHide)
                cubeFields[i].Orientation = self.pivotField['row']
        return self.pivotField['row']

    def addFieldToColumn(self, cubeFields, fieldName):
        numOfFields = len(cubeFields)
        for i in range(0, numOfFields):
            if cubeFields[i].Caption == fieldName:
                # print cubeFields[i].Caption + " : " + str(i) + " : " + str(cubeFields[1].DragToHide)
                cubeFields[i].Orientation = self.pivotField['column']
        return self.pivotField['column']

    def addFieldToFilter(self, cubeFields, fieldName):
        numOfFields = len(cubeFields)
        for i in range(0, numOfFields):
            if cubeFields[i].Caption == fieldName:
                # print cubeFields[i].Caption + " : " + str(i) + " : " + str(cubeFields[1].DragToHide)
                cubeFields[i].Orientation = self.pivotField['filter']
        return self.pivotField['filter']

    def addFieldToValue(self, cubeFields, fieldName):
        numOfFields = len(cubeFields)
        for i in range(0, numOfFields):
            if cubeFields[i].Caption == fieldName:
                # print cubeFields[i].Caption + " : " + str(i) + " : " + str(cubeFields[1].DragToHide)
                cubeFields[i].Orientation = self.pivotField['values']
        return self.pivotField['values']

    # def getTextField(self, pivotTable, text):
    #     print pivotTable.PivotFields().Count
    #     print pivotTable.PivotFields()[0].PivotFilters.Add2(17,None,None,None,None,"Duration")
    #     print cubeFields[0].PivotFields[0].Item("Duration")
    #     return

    def writeLog(self, message):
        # Open a file
        timeStamp = datetime.datetime.utcnow()

        fo = open("ExcelRegression.log", "a")
        fo.write("[" + str(timeStamp) + "] " + message + "\n");

        # Close opend file
        fo.close()