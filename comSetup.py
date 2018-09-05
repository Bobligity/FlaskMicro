import unittest
from ExcelCom import *


class comSetUp(unittest.TestCase):
    visibleTests = True
    com = ExcelCom()

    def setUp(self):
        # self.com = ExcelCom()
        self.com.setVisible(self.visibleTests)

    # def tearDown(self):
    #     if not self.visibleTests:
    #         self.com.setVisible(False)
    #         del self.com

    def testFramework(self):
        self.assertEqual(0, 0)

    def testComObject(self):
        self.assertNotEqual(None, self.com)

    def testOpenWorkbook(self):
        self.assertIsNotNone(self.com.openExistingWorkbook())

    def testGetDefaultVisibility(self):
        self.assertEqual(self.visibleTests, self.com.office.Visible)

    def testSetVisibility(self):
        self.com.setVisible(True)
        self.assertEqual(1, self.com.office.Visible)

    def testWorkBookField(self):
        self.assertIsNotNone(self.com.wb)

    def testGetWorksheetCount(self):
        self.assertNotEqual(0, self.com.getWorksheetCount())

    def testGetWorksheetList(self):
        self.assertIsNotNone(self.com.getWorksheets(self.com.wb))

    def testGetWorksheet(self):
        self.assertNotEqual(0, self.com.getWorksheet(0))

    def testGetPivotTables(self):
        self.assertIsNotNone(self.com.getPivotTables(self.com.getWorksheet(0)))

    def testGetPivotTable(self):
        self.assertIsNotNone(self.com.getPivotTable(0, self.com.getWorksheet(0)))

    def testGetPivotTableCount(self):
        self.assertIsNotNone(self.com.getPivotTableCount(self.com.getWorksheet(0)))

    def testGetPivotTableActiveFields(self):
        self.assertIsNotNone(self.com.getActivePivotTableFields(self.com.getPivotTable(0, self.com.getWorksheet(0))))

    def testGetAllFields(self):
        self.assertIsNotNone(self.com.getAllPivotTableFields(self.com.getPivotTable(0, self.com.getWorksheet(0))))

    def testRemoveField(self):
        before = self.com.getActivePivotTableFields(self.com.getPivotTable(0, self.com.getWorksheet(0))).Count
        orientation = self.com.removeField(
            self.com.getAllPivotTableFields(self.com.getPivotTable(0, self.com.getWorksheet(0))), "Title")
        after = self.com.getActivePivotTableFields(self.com.getPivotTable(0, self.com.getWorksheet(0))).Count
        # print str(before) + " > " + str(after)
        self.assertLess(after, before)
        self.assertEqual(0, orientation)

    def testAddFieldToRows(self):
        orientation = self.com.addFieldToRow(
            self.com.getAllPivotTableFields(self.com.getPivotTable(0, self.com.getWorksheet(0))), "Duration")
        self.assertEqual(1, orientation)

    def testAddFieldToColumns(self):
        orientation = self.com.addFieldToColumn(
            self.com.getAllPivotTableFields(self.com.getPivotTable(0, self.com.getWorksheet(0))), "User")
        self.assertEqual(2, orientation)

    def testAddFieldToFilters(self):
        orientation = self.com.addFieldToFilter(
        self.com.getAllPivotTableFields(self.com.getPivotTable(0, self.com.getWorksheet(0))), "Execution")
        self.assertEqual(3, orientation)


def testAddFieldToValues(self):
    orientation = self.com.addFieldToValue(
        self.com.getAllPivotTableFields(self.com.getPivotTable(0, self.com.getWorksheet(0))), "Fastest Success")
    self.assertEqual(4, orientation)


#def testDeferred(self):
#    @unittest.skip


def testAddTextToFilter(self):
    self.assertEqual("", self.com.getTextField(self.com.getPivotTable(0, self.com.getWorksheet(0)), ""))

    # def app(unittest.TestCase):