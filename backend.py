import sqlite3
from os import getcwd, listdir
import xlsxwriter

"""
(0, 'kod', 'TEXT'), 
(1, 'oglad', 'INTEGER'), 
(2, 'datezakl', 'TEXT'), 
(3, 'famil', 'TEXT'), 
(4, 'name', 'TEXT'), 
(5, 'tato', 'TEXT'), 
(6, 'pol', 'INTEGER'), 
(7, 'rod', 'TEXT'), 
(8, 'kodsity', 'INTEGER'), 
(9, 'kodraion', 'INTEGER'), 
(10, 'selo', 'TEXT'),      
(11, 'street', 'TEXT'), 
(12, 'woker', 'INTEGER'),
(13, 'social', 'INTEGER'), 
(14, 'special', 'TEXT'), # це професія з фахом
(15, 'mesto', 'TEXT'), # це освіта
(16, 'ministr', 'TEXT'), Це diagnoz_1
(17, 'organ', 'TEXT'), 
(18, 'lpu', 'TEXT'),
(19, 'moglad', 'INTEGER'), 
(20, 'grupinv', 'INTEGER'),
(21, 'toglad', 'INTEGER'), 
(22, 'roglad', 'INTEGER'), 
(23, 'invalid', 'INTEGER'), 
(24, 'poteri', 'INTEGER'), 
(25, 'prichin', 'INTEGER'), 
(26, 'faktor', 'TEXT'), 
(27, 'diagnoz', 'TEXT'), 
(28, 'expert', 'INTEGER'),
(29, 'lik', 'INTEGER'), 
(30, 'rebmed', 'TEXT'), НЕМА В НАКАЗІ
(31, 'rebwork', 'TEXT'), 
(32, 'rebhelp', 'TEXT'), СОЦІАЛЬНА ДОПОМОГА
(33, 'prog1', 'TEXT'), 
(34, 'prog2', 'TEXT'), 
(35, 'pr', 'INTEGER'), 
(36, 'prog3', 'TEXT'), 
(37, 'poterist', 'TEXT'), 
(38, 'prre', 'TEXT'), 
(39, 'rogladi', 'INTEGER'), РЕЗУЛЬТАТ ОГЛЯДУ 
(40, 'tran', 'TEXT'),
(41, 'grupinvi', 'TEXT') ПОПЕРЕДНЯ ГРУПА
"""

class DB:

	def __init__(self):
		pass


	def append_talon(self, num_msek, *args):
		if num_msek != 0 and args[0] != '':	#args[0] - це КОД ТАЛОНУ
			self.path = getcwd() + '/db_files/'
			self.conn = sqlite3.connect(self.path + 'db_{}.db'.format(num_msek))
			self.cur = self.conn.cursor()
			self.cur.execute('''INSERT INTO db(kod, oglad, datezakl, famil, name, tato, pol, rod, kodsity, kodraion, 
												selo, street, woker, social, special, mesto, ministr, organ, lpu,
												moglad, grupinv, grupinvi, toglad, roglad, rogladi, invalid, poteri, prichin, faktor, diagnoz, expert,
												rebwork, rebhelp, prog2, prog3)
								VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', args)
			self.conn.commit()
			self.cur.close()
			self.conn.close()
			print('ДОДАНО')
		else:
			print('НЕ ЗАПОВНЕНІ КОНТРОЛЬНІ ПОЛЯ! (НОМЕР КОМІСІЇ І КОД ТАЛОНУ)')


	def find_old_talon(self, message_var, num_msek, kod_var, old_kod_var, year_pop_tal_var, oglad_var, *args):
		self.path = getcwd() + '/db_files/'
		self.conn = sqlite3.connect(self.path + 'db_{}.db'.format(num_msek))
		self.cur = self.conn.cursor()
		try:
			if num_msek != 0 and kod_var != '' and old_kod_var != '' and year_pop_tal_var != '':
				self.find = self.cur.execute('''SELECT * FROM db WHERE kod LIKE ? AND datezakl LIKE ?''', [old_kod_var+'%', year_pop_tal_var+'%'])
				self.row = self.find.fetchone()
				args[0].set(self.row[6])                #   СТАТЬ
				args[1].set(self.row[3].strip(' '))     #   ПРІЗВ.
				args[2].set(self.row[4].strip(' '))     #   ІМЯ.
				args[3].set(self.row[5].strip(' '))     #   ПО БАТЬКОВІ
				args[4].set(self.row[7].split('-')[2])                #   ДАТА НАР день
				args[5].set(self.row[7].split('-')[1])                #   ДАТА НАР місяць
				args[6].set(self.row[7].split('-')[0])                #   ДАТА НАР рік
				args[7].set(self.row[8])               	#   МІСТО 
				args[8].set(self.row[9])               	#   РАЙОН
				args[9].set(self.row[10])              	#   СЕЛО
				args[10].set(self.row[11])              #   ВУЛИЦЯ
				args[11].set(self.row[12])              #   ПРАЦЮЄ.НЕПРАЦЮЄ
				args[12].set(self.row[13])              #   СОЦІАЛЬНА КАТЕГОРІЯ
				args[13].set(self.row[14].strip(' '))   #   професія з фахом
				args[14].set(self.row[15])              #   ОСВІТА
				args[15].set(self.row[16].strip(' '))   #   ДІАГНОЗ ДО
				args[16].set(self.row[18].strip(' '))   #   ПОТЯЖЧЕННЯ
				args[17].set(self.row[17].strip(' '))   #   КИМ НАПРАВЛЕНО
				#___ДРУГА_ПОЛОВИНА___
				args[18].set(self.row[2].split('-')[2])               # дата огл. ДЕНЬ
				args[19].set(self.row[2].split('-')[1])               # дата огл. МІСЯЦЬ
				args[20].set(self.row[2].split('-')[0])               # дата огл. РІК
				args[21].set(self.row[19])              # місце огл.
				args[22].set(self.row[20])              # поп.група А
				args[23].set(self.row[41])              # поп.група Б
				args[24].set(self.row[21])              # мета
				args[25].set(self.row[22])              # встан.група А
				args[26].set(self.row[39])              # встан.група Б
				args[27].set(self.row[23])              # строк
				args[28].set(self.row[24])              # відсотки
				args[29].set(self.row[25])              # причина
				args[30].set(self.row[27].strip(' '))   # діагноз_2
				args[31].set(self.row[36])              # працевлаштування
				args[32].set(self.row[28])              # експерт.рішення
			elif num_msek != 0 and kod_var != '' and old_kod_var == '' and year_pop_tal_var != '':
				self.find = self.cur.execute('''SELECT * FROM db WHERE kod LIKE ? AND datezakl LIKE ?''', [kod_var + '%', year_pop_tal_var + '%'])
				self.row = self.find.fetchone()
				oglad_var.set(self.row[1])
				args[0].set(self.row[6])                #   СТАТЬ
				args[1].set(self.row[3].strip(' '))     #   ПРІЗВ.
				args[2].set(self.row[4].strip(' '))     #   ІМЯ.
				args[3].set(self.row[5].strip(' '))     #   ПО БАТЬКОВІ
				args[4].set(self.row[7].split('-')[2])                #   ДАТА НАР день
				args[5].set(self.row[7].split('-')[1])                #   ДАТА НАР місяць
				args[6].set(self.row[7].split('-')[0])                #   ДАТА НАР рік
				args[7].set(self.row[8])               	#   МІСТО 
				args[8].set(self.row[9])               	#   РАЙОН
				args[9].set(self.row[10])              	#   СЕЛО
				args[10].set(self.row[11])              #   ВУЛИЦЯ
				args[11].set(self.row[12])              #   ПРАЦЮЄ.НЕПРАЦЮЄ
				args[12].set(self.row[13])              #   СОЦІАЛЬНА КАТЕГОРІЯ
				args[13].set(self.row[14].strip(' '))   #   професія з фахом
				args[14].set(self.row[15])              #   ОСВІТА
				args[15].set(self.row[16].strip(' '))   #   ДІАГНОЗ ДО
				args[16].set(self.row[18].strip(' '))   #   ПОТЯЖЧЕННЯ
				args[17].set(self.row[17].strip(' '))   #   КИМ НАПРАВЛЕНО
				#___ДРУГА_ПОЛОВИНА___
				args[18].set(self.row[2].split('-')[2])               # дата огл. ДЕНЬ
				args[19].set(self.row[2].split('-')[1])               # дата огл. МІСЯЦЬ
				args[20].set(self.row[2].split('-')[0])               # дата огл. РІК
				args[21].set(self.row[19])              # місце огл.
				args[22].set(self.row[20])              # поп.група А
				args[23].set(self.row[41])              # поп.група Б
				args[24].set(self.row[21])              # мета
				args[25].set(self.row[22])              # встан.група А
				args[26].set(self.row[39])              # встан.група Б
				args[27].set(self.row[23])              # строк
				args[28].set(self.row[24])              # відсотки
				args[29].set(self.row[25])              # причина
				args[30].set(self.row[27].strip(' '))   # діагноз_2
				args[31].set(self.row[36])              # працевлаштування
				args[32].set(self.row[28])              # експерт.рішення
			else:
				print('НЕ ЗАПОВНЕНІ ПОТРІБНІ ПОЛЯ! (НОМ.КОМІС. або ОГЛЯД або ПОП.КОД або РІК ПОП.ТАЛОНУ)')
				message_var.set('НЕ ЗАПОВНЕНІ ПОТРІБНІ ПОЛЯ! (НОМ.КОМІС. або ОГЛЯД або ПОП.КОД або РІК ПОП.ТАЛОНУ)')
		except:	
			print('ВІДСУТНІЙ В БД')
			message_var.set('ВІДСУТНІЙ В БД')


	def edit_talon(self, num_msek, *args):
		self.path = getcwd() + '/db_files/'
		self.conn = sqlite3.connect(self.path + 'db_{}.db'.format(num_msek))
		self.cur = self.conn.cursor()
		self.edit = self.cur.execute('''UPDATE db SET kod=?, oglad=?, datezakl=?, famil=?, name=?, tato=?, pol=?, rod=?, kodsity=?, kodraion=?, 
													  selo=?, street=?, woker=?, social=?, special=?, mesto=?, ministr=?, organ=?, lpu=?, moglad=?, 
													  grupinv=?, grupinvi=?, toglad=?, roglad=?, rogladi=?, invalid=?, poteri=?, prichin=?, faktor=?, diagnoz=?, 
													  expert=?, rebwork=?, rebhelp=?, prog2=?, prog3=?
													  WHERE kod=? AND datezakl=?''', (args[0],args[1],args[2],args[3],args[4],args[5],args[6],args[7],args[8],args[9],
																					  args[10],args[11],args[12],args[13],args[14],args[15],args[16],args[17],args[18],args[19],
																					  args[20],args[21],args[22],args[23],args[24],args[25],args[26],args[27],args[28],args[29],
																					  args[30],args[31],args[32],args[33],args[34],
																					  args[0],args[2]))
		self.conn.commit()
		self.cur.close()
		self.conn.close()
		print('РЕДАГОВАНО')


	def finder(self, num_msek, tree, first_name, name, birth_date):

		[tree.delete(row) for row in tree.get_children()]

		if num_msek == 0:
			self.path = getcwd() + '/db_files/'
			self.dir_list = listdir(path = self.path)
			for file in self.dir_list:
				self.conn = sqlite3.connect('{}'.format(self.path + file))
				self.cur = self.conn.cursor()
				self.find = self.cur.execute('''SELECT kod, oglad, datezakl, famil, name, tato, pol, rod, kodsity, kodraion, 
												selo, street, woker, social, special, mesto, ministr, organ, lpu,
												moglad, grupinv, grupinvi, toglad, roglad, rogladi, invalid, poteri, prichin, faktor, diagnoz, expert,
												rebwork, rebhelp, prog2, prog3 FROM db WHERE famil LIKE ? AND name LIKE ? AND rod LIKE ?''', 
												[first_name+'%', name+'%', str(birth_date)+'%'])
				for row in self.find.fetchall():
					row = list(row)
					row.insert(0, file)
					tree.insert('', 'end', values = row)

		elif num_msek != 0:
			self.path = getcwd() + '/db_files/'
			self.conn = sqlite3.connect(self.path + 'db_{}.db'.format(num_msek))
			self.cur = self.conn.cursor()
			self.find = self.cur.execute('''SELECT kod, oglad, datezakl, famil, name, tato, pol, rod, kodsity, kodraion, 
											selo, street, woker, social, special, mesto, ministr, organ, lpu,
											moglad, grupinv, grupinvi, toglad, roglad, rogladi, invalid, poteri, prichin, faktor, diagnoz, expert,
											rebwork, rebhelp, prog2, prog3 FROM db WHERE famil LIKE ? AND name LIKE ? AND rod LIKE ?''', 
											[first_name+'%', name+'%', str(birth_date)+'%'])
			for row in self.find.fetchall():
				row = list(row)
				row.insert(0, num_msek)
				tree.insert('', 'end', values = row)


	def delete_talon(self, num_msek, kod_var, date_ogl):
		if num_msek == 0:
			pass
		elif num_msek != 0 and kod_var != '' and date_ogl != '':
				self.path = getcwd() + '/db_files/'
				self.conn = sqlite3.connect(self.path + 'db_{}.db'.format(num_msek))
				self.cur = self.conn.cursor()
				self.delete = self.cur.execute('''DELETE FROM db WHERE kod=? AND datezakl=?''', (kod_var, date_ogl)) 
				self.conn.commit()
				self.cur.close()
				self.conn.close()


	def report_region(self, num_msek, date_Z, date_PO, diagnoz_Z, diagnoz_PO):
		self.DICT_CITY = {1:'Івано-Франківськ', 3:'Болехів', 25:'Яремче', 5:'Бурштин'}
		self.DICT_FIRST = {}
		self.DICT_SECOND = {}
		if num_msek == 0:
			print("ЦЕ ЩЕ НЕ ГОТОВО!!!")
		elif num_msek != 0:
			self.path = getcwd() + '/db_files/'
			self.conn = sqlite3.connect(self.path + 'db_{}.db'.format(num_msek))
			self.cur = self.conn.cursor()
			self.KODRAION = 1
			

			while self.KODRAION != 15:
				self.first_part=(
				#																		ВСЬОГО ВИЗНАНИХ ВПЕРШЕ
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad BETWEEN 0 and 3""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad BETWEEN 0 and 1""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 0""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 1""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 2""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 3""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 39 AND oglad = 1 AND
												datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 3""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 40 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 3""", [date_Z, date_PO, self.KODRAION]).fetchone()],
				#																		В Т.Ч. ПРАЦЮЮЧИХ
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad BETWEEN 0 and 3""", [date_Z, date_PO,  self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad BETWEEN 0 and 1""", [date_Z, date_PO,  self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 0""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 1""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 2""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 3""", [date_Z, date_PO, self.KODRAION]).fetchone()],
				#																		В Т.Ч. ПРАЦ. ВІКУ
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 3""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 1""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 0""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 1""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 2""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 3""", [date_Z, date_PO, self.KODRAION]).fetchone()],
							)
				self.DICT_FIRST[self.KODRAION] = self.first_part
				self.KODRAION += 1

			

			for key in self.DICT_CITY:
				self.second_part=(
				#																		ВСЬОГО ВИЗНАНИХ ВПЕРШЕ
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad BETWEEN 0 and 3""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad BETWEEN 0 and 1""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 0""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 1""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 2""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 3""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 39 AND oglad = 1 AND
												datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 3""", [date_Z, date_PO, self.KODRAION]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 40 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 3""", [date_Z, date_PO, self.KODRAION]).fetchone()],
				#																		В Т.Ч. ПРАЦЮЮЧИХ
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad BETWEEN 0 and 3""", [date_Z, date_PO,  key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad BETWEEN 0 and 1""", [date_Z, date_PO,  key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 0""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 1""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 2""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 3""", [date_Z, date_PO, key]).fetchone()],
				#																		В Т.Ч. ПРАЦ. ВІКУ
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 3""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 1""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 0""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 1""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 2""", [date_Z, date_PO, key]).fetchone()],
					[row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
												datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 3""", [date_Z, date_PO, key]).fetchone()],
							)		
				self.DICT_SECOND[key] = self.second_part
			self.write_to_xlsx(self.DICT_FIRST, self.DICT_SECOND)


	def write_to_xlsx(self, first_part, second_part):
		print(first_part, second_part)
		row = 0
		col = 0
		workbook = xlsxwriter.Workbook('report.xlsx')
		worksheet = workbook.add_worksheet('ЗВІТ')


if __name__ == "__main__":

	db = DB()