import pandas as pd

# need sqlalchemy to run pyodbc and pandas together
#   (https://stackoverflow.com/questions/71082494/getting-a-warning-when-using-a-pyodbc-connection-object-with-pandas)
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from sqlalchemy.sql import text
from sqlalchemy.orm import sessionmaker
from sqlalchemy.exc import ResourceClosedError
import sql_metadata

# LOGGER
import logging
logger = logging.getLogger()
if not(logger.hasHandlers()):
    stream_handler = logging.StreamHandler()
    stream_formatter = logging.Formatter('[%(levelname)s] %(message)s')
    stream_handler.setFormatter(stream_formatter)
    stream_handler.setLevel(logging.DEBUG)
    logger.addHandler(stream_handler)
    logger.setLevel(logging.DEBUG)


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
        engine = create_engine(conn_url, 
                               use_setinputsizes=False) ##https://github.com/sqlalchemy/sqlalchemy/issues/8681           
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
        

    # %% METHODS TO QUERY THE DB
    def execute_query(self, script: str) -> pd.DataFrame:
        '''
        Sends a query to the SQL engine to:
            - extract info
            - update/insert/delete
            
        If there is output, a dataframe will be returned.
        If not, it will return None.        
        '''

        with self.engine.begin() as conn:
            
            # executes the statement script
            result = conn.execute(text(script))

            # Depending on the type of query, there will be some or no output
            # when there is no output, there fetchall will throw an error    
            has_result = False
            try:            
                rows = result.fetchall()
                has_result = True
            except ResourceClosedError: #This result object does not return rows. It has been closed automatically.
                pass
            
            # If has result
            output = None
            if has_result:
                    
                columns = result.keys()
                df = pd.DataFrame(rows, columns=columns)
    
                # if data is empty, warn
                if df.empty:
                    # extract the source table name from the script
                    msg = f"The provided script returns an empty table as an output."
                    logger.debug(msg)
                    
                output = df
                
            # log
            output_str = "no output" if output is None \
                         else f"output of shape {str(output.shape)}"
            logger.debug(f"Completed query execution and got {output_str}.")

        return output
                
    # %% METHODS TO GET INFO FROM THE DB
    def get_all_tablenames(self):
        '''
        Get all tablenames from the connected db.
        
        Returns a list of tablenames.
        '''
        
        query = (
            f"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES "
            f"WHERE TABLE_CATALOG='{self.database}'"
            )
                
        df = self.execute_query(query)
        
        table_names = df["TABLE_NAME"].values   
        
        # when db gets updated with data, this will always be updated
        self.table_names = table_names
        
        return table_names
    
    def get_table_column_names(self, table_name):
        '''
        Get the column names for a table.

        Parameters
        ----------
        table_name : str

        Returns
        -------
        column_names : list
            List of all the column names
        '''
    
        query = (
            f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
            f"WHERE TABLE_NAME = N'{table_name}'"
            )
        
        df = self.execute_query(query)
        
        column_names = df["COLUMN_NAME"].values
        
        return column_names
    
    def count_table_nrows(self, table_name):
        
        query = f"SELECT COUNT(*) FROM {table_name}"
        
        df = self.execute_query(query)
        
        if df.shape != (1,1):
            err = "Unexpected output dimension."
            logger.error(err)
            raise Exception (err)
        
        #
        nrows = df.iat[0,0]
        
        return nrows
    
    def read_table(self,
                   table_name    = None,
                   query         = None,
                   column_subset = None):
        '''
        Reads a table from the db.

        Parameters
        ----------
        table_name : str, optional
            The name of the table to extract data from. 
            If this is not specified, query must be specified.
        query : str, optional
            The sql query to read the table. 
            If this is not specified, table_name must be specified.
        column_subset : list, optional
            Subset of columns to extract. 
            When it is not specified, the full table will be returned.
            
        Returns
        -------
        df : dataframe
        '''
                
        # Check that only table_name or query is specified
        # To get the final query
        if (table_name is None) and (query is None):
            err = (
                "Both table_name and query are not specified. "
                "Please specify either one.")
            logger.error(err)
            raise Exception (err)

        elif (table_name is not None) and (query is not None):
            
            err = (
                f"Both table_name ({table_name}) and "
                f"query ({query}) are specified. "
                "Please specify either one.")
            logger.error(err)
            raise Exception (err)
        
        elif table_name is not None:
            
            script = f"SELECT * FROM {table_name}"
            
        elif query is not None:
            
            script = query
            
            # Get the table name from the query
            table_name = ",".join(sql_metadata.Parser(query).tables)
        
        else:
            
            raise Exception ("Unusual situation.")
        
        # Query
        df = self.execute_query(script)
        
        # FIlter by column subset
        if column_subset is not None:
            
            df = df[column_subset]
            
        logger.debug(f"Completed reading the table '{table_name}'.")

        return df
    
    def read_all_tables(self):
        
        # Get all the table names
        table_names = self.get_all_tablenames()
        
        # Read all
        tablename_to_df = {tablename: self.read_table(tablename) for tablename in table_names}
        
        return tablename_to_df
    
    def export_all_tables(self, output_fp):
        
        # Load all the tables
        tablename_to_df = self.read_all_tables()
        
        # Create writer
        writer = pd.ExcelWriter(output_fp, engine='openpyxl')
        logger.debug("Writing tables to file={output_fp}.")
        for tablename, df in tablename_to_df.items():
            
            df.to_excel(writer, sheet_name = tablename)
            logger.debug(f"Table={tablename} written to file.")
        
        # Save
        writer.close()
        logger.debug(f"All tables saved to {output_fp}.")
        
    

    # %% METHODS TO WRITE TABLES TO SQL

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
            
            if str(e) == "__init__() got multiple values for argument 'schema'":
                msg += "\n\nThis could be due to incompatible versions of SQLAlchemy and pandas."
                msg += "\nSee also: https://stackoverflow.com/questions/75282511/df-to-table-throw-error-typeerror-init-got-multiple-values-for-argument"
            
            logger.error(msg)
            
            raise Exception(msg) #if  __init__() got multiple values for argument 'schema', version problem
                                 #https://stackoverflow.com/questions/75282511/df-to-table-throw-error-typeerror-init-got-multiple-values-for-argument

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
            
    def delete_table(self, table_name):
        '''
        Drop table.
        '''
        
        query = f'DROP TABLE {table_name};'
        self.execute_query(query)
        
        logger.info(f"Deleted table={table_name}.")
            
            
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

        # ---------------------------------------------------------------
        # prepare query 
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
        # ---------------------------------------------------------------
                
        # Delete
        nrow_before_del = self.count_table_nrows(table_name)
        
        try:
            
            self.execute_query(query)
            
            # Count #rows deleted
            nrow_after_del = self.count_table_nrows(table_name)
            ndeleted = nrow_before_del - nrow_after_del
            
            # Log
            logger.info(
                f"Successfully deleted {ndeleted} rows from {table_name}.")

        except:
            msg = f"Failed to delete rows from {table_name}."
            logger.error(msg)
            
    def create_table(self, table_name, structure_df):
        '''
        Parameters
        ----------
        tablename : TYPE
            DESCRIPTION.
        structure_df : df with two columns: COLUMN, TYPE.
            DESCRIPTION.
        '''
        
        # Check that all required columns are present
        required_columns = ["COLUMN", "TYPE"]
        missing_columns = set(required_columns).difference(structure_df.columns)
        if len(missing_columns) > 0:
            msg = "Missing columns: {missing_columns}."
            logger.error(msg)
            raise Exception (msg)
        
        # Prepare the script
        list_of_vartype = [
            f'"{s.at["COLUMN"]}" {s.at["TYPE"]}'for i, s in structure_df.iterrows()
            ]
        vartype_str = ", ".join(list_of_vartype)
        script = f"CREATE TABLE {table_name} ({vartype_str});"
        
        # Execute
        self.execute_query(script)
        logger.info(f"Created table={table_name} with {structure_df.shape[0]} columns and 0 rows.")


# %% FOR TESTING
if __name__ == "__main__":
    
    # Simple test of the class written in this file

    CONN_DICT = {
    "driver": "SQL Server",
    "server": r"U110\SQLDEV",
    "database": "fruits",
    }

    self = PyMsSQL(**CONN_DICT)
    
    self.execute_query('select * from dbo.color$')
    
    
    df = pd.DataFrame([[9, "green", 30],
                       [10, "purple", 20]], 
                      index =[0,1], 
                      columns=["id", "color", 'x'])
    
    # Test insert
    if False:

        self.insert_dataframe('color$', df)
        
        # Test delete
        self.delete('color$', ['id'], df)