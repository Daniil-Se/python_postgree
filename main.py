

import psycopg2
from psycopg2.extras import DictCursor #чтобы выводился список вместо кортежа


class postgree_connect:

	''' имя DB прописывать строкой '''

	def __init__(self, db_name):
		self.conn = psycopg2.connect(dbname = db_name, user = 'postgres', 
                        password = '2378951', host = 'localhost')
		self.cursor = self.conn.cursor(cursor_factory = DictCursor)
 
	def insert(self, table_name):
		self.cursor.execute("INSERT INTO {} VALUES ('ABC', '32213121') ".format(table_name)) 
		self.conn.commit() #сохраняем изменения в БД

	def show_table(self, table_name):
		self.cursor.execute("SELECT * FROM {}".format(table_name)) #LIMIT 2')
		for row in self.cursor:
			print(row)

	def close(self):		

		''' всегда закрывать каретку и БД'''

		self.cursor.close() #закрываем каретку
		self.conn.close() #закрываем БД

	def __str__(self):
		return 'class Object postgree'


a = postgree_connect('test_db')
a.show_table('test_table')
a.close()
