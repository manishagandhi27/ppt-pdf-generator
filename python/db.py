import mysql.connector


class MySQLClient():

    def __init__(self, host, user, password, database):
        self.db_connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database)


    def execute_query(self, query):
        try:
            db_cursor = self.db_connection.cursor()
            db_cursor.execute((query))
            result = db_cursor.fetchall()
            print ('Query -> ', query)
            print ('Result -> ', result)
            db_cursor.close()
            self.db_connection.close()
            return result
        except Exception as e:
            print ('Exception -> ', e)
