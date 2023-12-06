import pandas as pd

# need sqlalchemy to run pyodbc and pandas together
#   (https://stackoverflow.com/questions/71082494/getting-a-warning-when-using-a-pyodbc-connection-object-with-pandas)
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from sqlalchemy.sql import text
from sqlalchemy.orm import sessionmaker
from sqlalchemy.types import *
from sqlalchemy.exc import *

from settings import logger


class PyMsSQL:

    def __init__(self,
                 driver=None,
                 server=None,
                 database=None,
                 username=None,
                 password=None):
        '''
        This class establishes a connection to the MS SQL server
        and database.

        Parameters
        ----------
        driver : str
            The driver to connect to the server. 
            For example, "SQL Server" for MS SQL server.
            This must be set.

        server: str
            The name of the server to connect to.
            This must be set.

        database: str
            The name of database to connect to.
            This must be set.

        username: str

        password: str                    
        '''

        # Save as attributes (validation will be performed)
        self.driver = driver
        self.server = server
        self.database = database
        self.username = username  # optional
        self.password = password  # optional, if username is provided, password is required

        # Initialise the connection command
        conn_str = (
            f"Driver={self.driver};"
            f"Server={self.server};"
            f"Database={self.database};"
            f"Trusted_Connection=no;"
        )

        # If username provided, add username and password to conn_str
        if self.username:
            # Set the user name and password
            optional_str = (
                f"UID={self.username};"
                f"PWD={self.password};"
            )

            # Append
            conn_str = conn_str + optional_str

        # Connect
        logger.debug("Connecting to database...")

        conn_url = URL.create('mssql+pyodbc', query={'odbc_connect': conn_str})
        engine = create_engine(conn_url)
        connection = engine.raw_connection()
        cursor = connection.cursor()

        Session = sessionmaker(bind=engine)
        session = Session()

        logger.info(f"Connected to {self.database}.")

        self.conn_str = conn_str
        self.engine = engine
        self.cursor = cursor
        self.session = session

    @property
    def driver(self) -> str:
        return self._driver

    @driver.setter
    def driver(self, value):
        # Check if the given value is None
        if value is None:
            msg = f"The value of the key `driver` in CONN_DICT cannot be None. Edit in `settings.py`."
            logger.error(msg)
            raise ValueError(msg)

        # Check the type of the given value (required type: str)
        if not isinstance(value, str):
            msg = f"The value of the key `driver` in CONN_DICT should be str. Edit in `settings.py`."
            logger.error(msg)
            raise ValueError(msg)

        # Check the format of the given value (required format \{.*\})
        if not value.startswith("{"):
            value = "{" + value
        if not value.endswith("}"):
            value = value + "}"

        self._driver = value

    @property
    def server(self) -> str:
        return self._server

    @server.setter
    def server(self, value):
        # Check if the given value is None
        if value is None:
            msg = f"The value of the key `server` in CONN_DICT cannot be None. Edit in `settings.py`."
            logger.error(msg)
            raise ValueError(msg)

        self._server = value

    @property
    def database(self) -> str:
        return self._database

    @database.setter
    def database(self, value):
        # Check if the given value is None
        if value is None:
            msg = f"The value of the key `database` in CONN_DICT cannot be None. Edit in `settings.py`."
            logger.error(msg)
            raise ValueError(msg)

        self._database = value

    @property
    def username(self) -> str:
        return self._username

    @username.setter
    def username(self, value):
        self._username = value

    @property
    def password(self) -> str:
        return self._password

    @password.setter
    def password(self, value):
        if self.username is not None and value is None:
            msg = f"Username is provided but password is not set. Edit in `settings.py`."
            logger.error(msg)
            raise ValueError(msg)

        self._password = value

    # %% METHODS TO READ TABLES FROM SQL
    def read_sql(self, script: str) -> pd.DataFrame:

        with self.engine.begin() as conn:
            # executes the statement script to write to the database
            result = conn.execute(text(script))

            rows = result.fetchall()
            columns = result.keys()

            df = pd.DataFrame(rows, columns=columns)

            # if data is empty, warn
            if df.empty:
                # extract the source table name from the script
                msg = f"The provided script returns an empty table as an output."
                logger.debug(msg)

        return df

    def read_sql_full_table(self,
                            table_name: str,
                            read_current_only: bool = False,
                            start_column_name: str = None,
                            end_column_name: str = None
                            ) -> pd.DataFrame:
        '''
        Reads the entire contents of a specified table from a database.

        Parameters
        ----------
        table_name : str
            The name of the table to read from.
        read_current_only : bool
            Set to True to read only those rows that are currently active
        start_column_name : str
            Start column name to read date information. e.g. start_date
        end_column_name : str
            End column name to read date information. e.g. end_date

        Returns
        -------
        df : pd.DataFrame
            A pandas DataFrame containing the data from the specified
            table.
        '''
        if read_current_only:
            if (start_column_name is None) or (start_column_name is None):
                msg = ("The parameter 'start_column_name' or 'end_column_name' ",
                       "must be specified to filter the data. ",
                       "Reading the whole data...")
                logger.warning(msg)

            else:
                script = f'''SELECT * FROM {table_name}
                          WHERE GETDATE() BETWEEN {start_column_name} AND {end_column_name}'''

        else:
            script = f"SELECT * FROM {table_name}"

        df = self.read_sql(script)

        logger.debug(f"Completed reading the table '{table_name}'.")

        return df

    # %% METHODS TO WRITE TABLES TO SQL
    def execute_query(self, script: str):
        '''
        Execute the provided script.

        Parameters
        ----------
        script : str
            The SQL script to execute for data insertion or update.

        Raises
        ------
        Exception
            If there is an error executing the given script.
        '''

        logger.debug(f"Executing query \n'{script}'...")

        try:
            with self.engine.begin() as conn:
                # executes the statement script to write to the database
                conn.execute(text(script))

        except Exception as e:
            logger.error(e)
            raise e

        logger.debug(f"Completed query execution.")

    def insert_dataframe(self, table_name: str, df: pd.DataFrame):
        '''
        Inserts data from a pandas DataFrame into a specified SQL table.

        Parameters
        ----------
        table_name : str
            The name of the table to insert data into.
        df : pd.DataFrame
            The pandas DataFrame containing the data to be inserted.

        Raises
        ------
        Exception
            If there is an error inserting data into the database.

        '''

        try:
            df.to_sql(name=table_name,   # name of table
                      con=self.engine,  # connection engine
                      if_exists='append',     # what to do if table exists
                      index=False         # whether to use index
                      )

            logger.info(
                f"Successfully executed insertion of {df.shape[0]} row(s) into {table_name}.")

        except Exception as e:
            msg = (f"Failed to insert dataframe into {table_name}."
                   f"\n{e}")
            logger.error(msg)
            raise Exception(msg)

        logger.debug(f"Completed dataframe insertion into {table_name}.")

    def update(self, table_name: str, id_cols: list, data: dict):
        '''
        Updates records in a specified SQL table based on the provided
        DataFrame.

        Parameters
        ----------
        table_name : str
            The name of the table to update.
        id_cols : list
            The list of column names used for record identification.
            This id_col will be used to identify the row to update in the
            WHERE clause in the UPDATE SQL statement.
        data : dict
            data containing the updated data.

        Raises
        ------
        Exception
            If the specified id_col cannot be found in data.
        Exception
            If there is an error updating records in the database.
        '''

        if not isinstance(id_cols, list):
            msg = (f"Identifier column name(s) '{id_cols}' provided is "
                   f"in {type(id_cols)} data type. "
                   f"Expected <class 'list'>.")
            logger.error(msg)
            raise Exception(msg)

        id_cols_set = set(id_cols)

        # to check if listed id_cols exists in the provided dict
        if not id_cols_set.issubset(set(data.keys())):
            msg = (f"The given identification column(s) does not exist: "
                   f"{id_cols_set}")
            logger.error(msg)
            raise Exception(msg)

        # structure of an UPDATE script:
        #   '''
        #   UPDATE table_name
        #   SET col_1 = val1, col_2 = val2, ...
        #   WHERE id_col = id1
        #   '''

        query_lst = []

        df = pd.DataFrame(data)

        for i in range(len(df)):
            # loop through the number of rows in df
            #   each row will be updated individually, with a statement
            #   script starting with 'UPDATE'

            df_with_updates = df.drop(columns=id_cols)
            cols_to_update = list(df_with_updates.columns)

            conditions = []
            for id_col in id_cols:
                conditions.append(f"{id_col} = {df[id_col].iloc[i]}")

            conditions = (" AND ").join(conditions)

            updates = []
            for j in range(len(cols_to_update)):
                updates.append(
                    f"{cols_to_update[j]} = '{df_with_updates.iloc[i,j]}'")

            updates = (", ").join(updates)

            update_statement = f"UPDATE {table_name} SET {updates} WHERE {conditions};"

            query_lst.append(update_statement)

        # If query is longer than 300 statements, do it repeatedly instead of doing at once
        #  Otherwise, sometimes update query silently fails
        rep = len(query_lst) // 300
        remainder = len(query_lst) % 300

        query_lst_by_300 = []
        if rep > 0:
            for i in range(rep):
                query_temp = query_lst[i*300:(i+1)*300]
                query_lst_by_300.append("\n".join(query_temp))

        query_lst_by_300.append("\n".join(query_lst[-remainder:]))

        try:
            for query in query_lst_by_300:
                self.execute_query(query)
                len_of_query = len(query.split("\n"))
                logger.info(
                    f"Successfully updated {len_of_query} row(s) in {table_name}.")

        except:
            msg = f"Failed to update {table_name}."
            logger.error(msg)

    def delete(self, table_name: str, id_cols: list, df: pd.DataFrame):
        '''
        Writes SQL DELETE statements to delete rows from a table based
        on given conditions.

        Parameters
        ----------
        table_name : str
            The name of the table from which rows will be deleted.
        id_cols : list
            A list of column names that uniquely identify the rows to be deleted.
        df : pd.DataFrame
            The DataFrame containing the data to be deleted.
        '''
        id_cols_set = set(id_cols)

        # to restrict deletion of whole table
        if not id_cols:
            msg = ("Please specify id_cols to be used as condition to delete rows.")
            logger.error(msg)
            raise Exception(msg)

        # to check if listed id_cols exists in the provided dataframe
        if not id_cols_set.issubset(df.columns):
            msg = (f"Column names {id_cols} provided cannot be found "
                   f"in provided dataframe.")
            logger.error(msg)
            raise Exception(msg)

        # structure of a DELETE script:
        #   '''
        #   DELETE FROM table_name
        #   WHERE condition
        #   '''

        query_lst = []

        for i in range(len(df)):
            # loop through the number of rows in df
            #   each row will be updated individually
            conditions = []
            for id_col in id_cols:
                conditions.append(f"{id_col} = {df[id_col].iloc[i]}")

            conditions = (" AND ").join(conditions)

            delete_statement = f"DELETE FROM {table_name} WHERE {conditions};"

            # append script_str to list of queries
            query_lst.append(delete_statement)

        query = "\n".join(query_lst)

        try:
            self.execute_query(query)

            logger.info(
                f"Successfully deleted {len(query_lst)} rows from {table_name}")

        except:
            msg = f"Failed to delete rows from {table_name}."
            logger.error(msg)


# %% FOR TESTING
if __name__ == "__main__":
    from settings import CONN_DICT

    self = PyMsSQL(**CONN_DICT)