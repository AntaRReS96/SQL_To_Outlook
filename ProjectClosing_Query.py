import time
import pandas as pd
import sqlalchemy as sa
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.engine import URL

class SQL:

    def __init__(self):
        # deklaracja zmiennych do połączenia z bazą danych
        self.server = ''
        self.database = ''
        self.username = ''
        self.password = ''
        
        # utworzenie silnika do połączenia z bazą danych
        self.db_url = URL.create(drivername='mssql+pyodbc', 
                                 username=self.username, 
                                 password=self.password, 
                                 host=self.server, 
                                 database=self.database, 
                                 query={'driver': 'SQL Server'})
        self.engine = sa.create_engine(self.db_url)
        
        # utworzenie zapytania do bazy danych
        self.query1 = """
            SELECT *
            FROM <data_source>
            """


    def get_data(self):
        # pobranie danych z SQL i zwrócenie Data Frame
        self.engine.connect()
        df = pd.read_sql_query(self.query1, self.engine)
        self.engine.dispose()
        return df
    
    def get_data_from_SQL_to_DataFrame(self):
        max_attempts = 3
        attempts = 0
        while attempts < max_attempts:
            try:
                self.df = self.get_data()
                return self.df
            except SQLAlchemyError as ex:
                if "deadlock victim" in str(ex):
                    print(f"Deadlock detected. Retrying in 10 seconds. Attempt {attempts+1} of {max_attempts}")
                    attempts += 1
                    time.sleep(10)
                else:
                    raise ex
    
    def get_data_from_SQL_to_excel(self):
        max_attempts = 3
        attempts = 0
        while attempts < max_attempts:
            try:
                self.df = self.get_data()
                self.df.to_excel('Sample_Query.xlsx', index=False)
                break
            except SQLAlchemyError as ex:
                if "deadlock victim" in str(ex):
                    print(f"Deadlock detected. Retrying in 10 seconds. Attempt {attempts+1} of {max_attempts}")
                    attempts += 1
                    time.sleep(10)
                else:
                    raise ex
