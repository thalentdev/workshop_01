# IMPORT THE SQALCHEMY LIBRARY's CREATE_ENGINE METHOD
from sqlalchemy import create_engine

# DEFINE THE DATABASE CREDENTIALS
user = 'app_user'
password = 'password123'
host = '192.168.0.221'
port = 18432
database = 'Adventureworks'


# PYTHON FUNCTION TO CONNECT TO THE MYSQL DATABASE AND
# RETURN THE SQLACHEMY ENGINE OBJECT
def get_connection():
    dbschema = 'sales,procurement'
    return create_engine(
        url="postgresql+psycopg2://{0}:{1}@{2}:{3}/{4}".format(
            user, password, host, port, database
        ), connect_args={'options': '-csearch_path={}'.format(dbschema)}
    )

