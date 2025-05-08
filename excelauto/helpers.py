import os

import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, Reference, Series

from db import get_connection


def add_tabular_data():
    db_engine = get_connection()
    df = pd.read_sql('SELECT s.name, s.salesytd, s.saleslastyear FROM sales.salesterritory s', db_engine)
    path = f'./excelauto/reports'

    # Check whether the specified path exists or not
    is_exist = os.path.exists(path)
    if not is_exist:
        # Create a new directory because it does not exist
        os.makedirs(path)

    writer = pd.ExcelWriter(f'{path}/sales.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sales Data')
    writer.close()


def add_chart():
    path = f'./excelauto/reports/sales.xlsx'
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    chart = BarChart()
    chart.type = "col"
    chart.title = "Sales Summary"
    chart.y_axis.title = 'Sales'
    chart.x_axis.title = 'Area Name'
    chart.height = 10  # default is 7.5
    chart.width = 20  # default is 15

    categories = Reference(ws, min_col=1, max_col=1, min_row=2, max_row=11)

    values_1 = Reference(ws, min_col=2, max_col=2, min_row=2, max_row=11)
    series_1 = Series(values_1, title='Sales Ytd')
    chart.append(series_1)

    values_2 = Reference(ws, min_col=3, max_col=3, min_row=2, max_row=11)
    series_2 = Series(values_2, title='Sales Last Year')
    chart.append(series_2)

    chart.set_categories(categories)

    ws.add_chart(chart, "E3")
    wb.save(path)
