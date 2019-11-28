
import openpyxl as op
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


class open_xlsx:

	def __init__(self):
		self.wb = op.load_workbook(filename = 'd:/Users/A/Desktop/тест.xlsx')
		self.sheets = self.wb.sheetnames
		self.ws1 = self.wb[self.sheets[0]]
		self.ws1_A_column = self.ws1['A']
		for row in self.ws1_A_column[1:]:
			print(row.value)

#a = postgree_connect('test_db')
#a.show_table('test_table')
#a.close()

op_xl = open_xlsx()