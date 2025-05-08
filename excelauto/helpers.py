import os
from db import get_connection
import pandas as pd

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