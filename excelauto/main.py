from openpyxl.chart.axis import Scaling
from openpyxl.chart.legend import Legend
from openpyxl.workbook.workbook import Workbook

from db import get_connection
from sqlalchemy import text, Engine
import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.styles import Font
import pandas as pd
from helpers import add_tabular_data


def main():
    # db_engine = get_connection()
    # sql = text('SELECT * FROM sales.salesterritory')
    # with db_engine.connect() as conn:
    #     result = conn.execute(sql).fetchall()

    add_tabular_data()

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
    # data = Reference(ws, min_col=2, max_col=3, min_row=2, max_row=11)
    # chart.legend = None
    # chart.add_data(data)

    values_1 = Reference(ws, min_col=2, max_col=2, min_row=2, max_row=11)
    series_1 = Series(values_1, title='Sales Ytd')
    chart.append(series_1)

    values_2 = Reference(ws, min_col=3, max_col=3, min_row=2, max_row=11)
    series_2 = Series(values_2, title='Sales Last Year')
    chart.append(series_2)

    chart.set_categories(categories)

    # series1_values = Reference(sheet, min_col=2, min_row=2, max_col=2, max_row=11)
    # series2_values = Reference(sheet, min_col=3, min_row=2, max_col=3, max_row=11)

    # chart.title = "Sales YTD"
    # chart.x_axis.title = "Name"
    # chart.y_axis.title = "Sales in USD"

    ws.add_chart(chart, "E3")
    wb.save(path)

    print('DONE')


# if directly on app.py not as module, will run below script
if __name__ == '__main__':
    main()
