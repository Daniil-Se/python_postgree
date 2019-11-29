
import openpyxl as op
import psycopg2
from psycopg2.extras import DictCursor #чтобы выводился список вместо кортежа


class postgree_connect:

	''' При создании экземпляра прописыватьимя бд для подключения'''
	''' имя DB прописывать строкой '''

	def __init__(self, db_name):
		self.conn = psycopg2.connect(dbname = db_name, user = 'postgres', 
                        password = '2378951', host = 'localhost')
		self.cursor = self.conn.cursor(cursor_factory = DictCursor)
 
	def insert(self, table_name, values):
		''' Для инсерта прописывать имя таблицы строкой'''
		''' Values - передавать кортежем '''

		#self.cursor.execute("INSERT INTO {} VALUES ('ABC', '32213121') ".format(table_name)) 
		self.cursor.execute("INSERT INTO {} VALUES {}".format(table_name, values)) 
		self.conn.commit() #сохраняем изменения в БД

	def show_table(self, table_name):
		''' Для показа всех данных прописывать имя таблицы строкой'''

		self.cursor.execute("SELECT * FROM {}".format(table_name)) 
		for row in self.cursor:
			print(row)

	def close(self):		
		''' всегда закрывать каретку и БД'''

		self.cursor.close() #закрываем каретку/курсор
		self.conn.close() #закрываем БД


	def __str__(self):
		return 'class Object postgree DB'



class open_xlsx:

	def __init__(self, file_name, sheet_number):
		''' Путь к файлу передавать строкой '''
		''' Номер листа для копирования передавать целочисленным'''
		self.wb = op.load_workbook(filename = file_name)
		self.sheets = self.wb.sheetnames #загружаем список листов

		self.ws = self.wb[self.sheets[sheet_number-1]] #передаем номер листа
		''' Здесь обращение происходит по индексу, поэтому -> минус один'''

				

	def take_the_values(self, name_column_end, num_of_line):
		''' Имя колонки писать строкой, этот аргумент определяет конец считывания '''
		''' Выводит одну строку под определенным номером '''

		insert_spisok = []
		for column in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 
					   		'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
					   		'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG'     				   ]:
						
			insert_spisok.append(self.ws[column][num_of_line].value) 

			if column == name_column_end: # условие отвечающее за выход из цикла, если цикл дошел до нуной нам буквы 
				break

		return tuple(insert_spisok)


op_xl_1 = open_xlsx('d:/Users/A/Desktop/тест.xlsx', 1)

a = postgree_connect('test_db')
a.insert('experts.expert_test', op_xl_1.take_the_values('B', 1))
a.show_table('experts.expert_test')
a.close()

