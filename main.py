from ExcelCom import *
# import comSetup
import win32com.client
from flask import Flask, request
import pandas as pd
from bokeh.plotting import figure, output_file, show, save

pd.set_option('expand_frame_repr', False)

office = win32com.client.Dispatch("Excel.Application")
wb = office.Workbooks.Open(r"C:\Users\Nikita\Documents\Josh\VR_pivots\sales.xlsx")

# fact = csv, dimension = csv
# load wb that has blank pivot table
# save as diff file


# Flask application
app = Flask(__name__)

@app.route("/")
def hello():
    return "Hello World!"

@app.route('/facts')
def facts():
    # get facts from URL using query strings
    facts = request.args.get('facts').split(',')

@app.route('/dimensions')
def dimensions():
    # get dimensions from URL using query strings
    dimensions = request.args.get('dimensions').split(',')

@app.route('/getFields')
def fields():
    # Getting fields from pivot table
    ws = wb.Worksheets
    # We know pivot table is on worksheet[1]
    active_ws = ws[0]
    pvt = active_ws.PivotTables()
    active_pvt = pvt[0]

    list_active_wb_flds = []
    def getActivePivotTableFields(table):
        # Returns a list of fields minus values which should be the last column
        numOfFields = len(table.PivotFields())
        for i in range (0, numOfFields):
            list_active_wb_flds.append(table.PivotFields()[i].Caption)
        return list_active_wb_flds[:-1]

    active_wb_flds = getActivePivotTableFields(active_pvt)
    return active_wb_flds

@app.route('/serveData')
# Get data from pivot table. Return to Unity app.
def getflaskdata():
    facts = request.args.get('facts')
    dimensions = request.args.get('dimensions')
    office = win32com.client.Dispatch("Excel.Application")
    # self.setVisible(True)
    wb = office.Workbooks.Open(r"C:\Users\Nikita\Documents\Josh\VR_pivots\sales.xlsx")

    # Remove all fields
    orientation = wb.removeField(
        wb.getAllPivotTableFields(wb.getPivotTable(0, wb.getWorksheet(0))), "Title")

    # Add fields that are desired
    for fact in facts:
        orientation = wb.addFieldToRow(
            wb.getAllPivotTableFields(wb.getPivotTable(0, wb.getWorksheet(0))), fact)
    for dim in dimensions:
        orientation = wb.addFieldToValue(
            wb.getAllPivotTableFields(wb.getPivotTable(0, wb.getWorksheet(0))), dim)


    df = pd.read_excel(r"C:\Users\Nikita\Documents\Josh\VR_pivots\sales.xlsx")
    df.drop("Unnamed: 0", axis=1, inplace=True)

# TODO: Plot based on what is fed to us and send to Unity to render.

# output to static HTML file
    output_file("lines.html")

    # create a new plot with a title and axis labels
    p = figure(title="simple line example", x_axis_label='x', y_axis_label='y')

    x = df.Index.values
    y = df["Average of Fuel_Price"]
    # add a line renderer with legend and line thickness
    p.line(x, y, line_width=2)

    show(p)


# Trying out bokeh and graphing

df = pd.read_excel(r"C:\Users\Nikita\Documents\Josh\VR_pivots\sales.xlsx")
df.drop("Unnamed: 0", axis=1, inplace=True)

# output to static HTML file
output_file("lines.html")

# create a new plot with a title and axis labels
p = figure(title="simple line example", x_axis_type='datetime', x_axis_label='x', y_axis_label='y')

x = df['Date']
y = df["Average of Fuel_Price"]
# add a line renderer with legend and line thickness
p.circle(x, y, line_width=2)

show(p)