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
        if num_msek != 0 and args[0] != '': #args[0] - це КОД ТАЛОНУ
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
                args[7].set(self.row[8])                #   МІСТО 
                args[8].set(self.row[9])                #   РАЙОН
                args[9].set(self.row[10])               #   СЕЛО
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
                args[7].set(self.row[8])                #   МІСТО 
                args[8].set(self.row[9])                #   РАЙОН
                args[9].set(self.row[10])               #   СЕЛО
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
        self.DICT_CITY = {1:'Івано-Франківськ', 
                          3:'Болехів', 
                          25:'Яремче', 
                          5:'Бурштин'}

        self.DICT_RAION = {1:'Богородчанський', 2:'Верховинський', 
                           3:'Галицький', 4:'Городенківський', 
                           5:'Долинський', 6:'Калуський',
                           7:'Коломийський', 8:'Косівський', 
                           9:'Надвірнянський', 10:'Рогатинський', 
                           11:'Рожнятівський', 12:'Снятинський', 
                           13:'Тисменицький', 14:'Тлумацький'}
        
        if num_msek == 0:
            
            self.DICT_FIRST = {1 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               2 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               3 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               4 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               5 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               6 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               7 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               8 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               9 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               10 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               11 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               12 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               13 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               14 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]}
            
            self.DICT_SECOND = {1 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                                3 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                                25 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                                5 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]}

            self.path = getcwd() + '/db_files/'
            self.dir_list = listdir(path = self.path)
            
            for db_name in self.dir_list:
                self.conn = sqlite3.connect(self.path + db_name)
                self.cur = self.conn.cursor()
                self.KODRAION = 1
                while self.KODRAION < 15:
                    self.first_part=(
                    #                                                                       ВСЬОГО ВИЗНАНИХ ВПЕРШЕ
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad BETWEEN 0 and 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad BETWEEN 0 and 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 39 AND oglad = 1 AND
                                                    datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 40 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    #                                                                       В Т.Ч. ПРАЦЮЮЧИХ
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad BETWEEN 0 and 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO,  self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad BETWEEN 0 and 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO,  self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    #                                                                       В Т.Ч. ПРАЦ. ВІКУ
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                                )
                    
                    self.first_list = []

                    for i in self.first_part:
                        self.first_list.append(i[0])

                    for i, v in enumerate(self.first_list):
                        self.DICT_FIRST[self.KODRAION].append(v + self.DICT_FIRST[self.KODRAION][i])
                        self.DICT_FIRST[self.KODRAION].pop(0)
                
                    self.KODRAION += 1
                
                for key in self.DICT_CITY:
                    self.second_part=(
                    #                                                                       ВСЬОГО ВИЗНАНИХ ВПЕРШЕ
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad BETWEEN 0 and 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad BETWEEN 0 and 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 39 AND oglad = 1 AND
                                                    datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 40 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    #                                                                       В Т.Ч. ПРАЦЮЮЧИХ
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad BETWEEN 0 and 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad BETWEEN 0 and 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    #                                                                       В Т.Ч. ПРАЦ. ВІКУ
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                        [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                    datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                                )       
                    
                    self.second_list = []

                    for i in self.second_part:
                        self.second_list.append(i[0])
                
            
                    for i, v in enumerate(self.second_list):
                        self.DICT_SECOND[key].append(v + self.DICT_SECOND[key][i])
                        self.DICT_SECOND[key].pop(0)
                
            #for K, V in self.DICT_FIRST.items():
            #    print(K, V)

            #for K, V  in self.DICT_SECOND.items():
            #    print(K, V)

            self.write_to_xlsx(self.DICT_CITY, self.DICT_RAION, self.DICT_FIRST, self.DICT_SECOND, num_msek, date_Z, date_PO, diagnoz_Z, diagnoz_PO)


        elif num_msek != 0:
            self.DICT_FIRST = {1 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               2 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               3 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               4 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               5 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               6 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               7 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               8 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               9 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               10 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               11 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               12 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               13 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                               14 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]}
            
            self.DICT_SECOND = {1 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                                3 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                                25 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 
                                5 : [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]}
            self.path = getcwd() + '/db_files/'
            self.conn = sqlite3.connect(self.path + 'db_{}.db'.format(num_msek))
            self.cur = self.conn.cursor()
            self.KODRAION = 1
            
            while self.KODRAION < 15:
                self.first_part=(
                #                                                                       ВСЬОГО ВИЗНАНИХ ВПЕРШЕ
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad BETWEEN 0 and 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad BETWEEN 0 and 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 39 AND oglad = 1 AND
                                                datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 40 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                #                                                                       В Т.Ч. ПРАЦЮЮЧИХ
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad BETWEEN 0 and 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO,  self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad BETWEEN 0 and 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO,  self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and woker = 1 AND roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                #                                                                       В Т.Ч. ПРАЦ. ВІКУ
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad BETWEEN 0 AND 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodraion = ? AND roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, self.KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()],
                            )
                
                self.first_list = []

                for i in self.first_part:
                    self.first_list.append(i[0])

                for i, v in enumerate(self.first_list):
                    self.DICT_FIRST[self.KODRAION].append(v + self.DICT_FIRST[self.KODRAION][i])
                    self.DICT_FIRST[self.KODRAION].pop(0)
            
                self.KODRAION += 1
            

            for key in self.DICT_CITY:
                self.second_part=(
                #                                                                       ВСЬОГО ВИЗНАНИХ ВПЕРШЕ
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad BETWEEN 0 and 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad BETWEEN 0 and 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 39 AND oglad = 1 AND
                                                datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 40 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                #                                                                       В Т.Ч. ПРАЦЮЮЧИХ
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad BETWEEN 0 and 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad BETWEEN 0 and 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodsity = ? and woker = 1 AND roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                #                                                                       В Т.Ч. ПРАЦ. ВІКУ
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad BETWEEN 0 AND 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 1 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 0 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 2 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                    [row for row in self.cur.execute("""SELECT COUNT(roglad) FROM db WHERE strftime('%Y', 'now') - rod >= 18 AND strftime('%Y', 'now') - rod <= 60 AND oglad = 1 AND 
                                                datezakl BETWEEN ? AND ? AND kodsity = ? AND roglad = 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, key, diagnoz_Z, diagnoz_PO]).fetchone()],
                            )

                self.second_list = []

                for i in self.second_part:
                    self.second_list.append(i[0])
            
                for i, v in enumerate(self.second_list):
                    self.DICT_SECOND[key].append(v + self.DICT_SECOND[key][i])
                    self.DICT_SECOND[key].pop(0)

            #for K, V in self.DICT_FIRST.items():
            #    print(K, V)

            #for K, V  in self.DICT_SECOND.items():
            #    print(K, V)

            self.write_to_xlsx(self.DICT_CITY, self.DICT_RAION, self.DICT_FIRST, self.DICT_SECOND, num_msek, date_Z, date_PO, diagnoz_Z, diagnoz_PO)


    def write_to_xlsx(self, DICT_CITY, DICT_RAION, first_part, second_part, num_msek, date_Z, date_PO, diagnoz_Z, diagnoz_PO):
        #print(first_part, second_part)
        self.DICT_KOMIS = {0:'ВСІ МСЕК', 1:'Обласна', 2:'Міськміжрайонна', 
                           3:'Калуська', 4:'Коломийська', 5:'Кардіологічна', 
                           6:'Міська', 7:'Травматологічна', 8:'Психіатрична', 
                           9:'Фтизіо-пульмонологічна', 10:'Міжрайонна', 11:'Коломия №2'}
        row = 7
        col = 0
        workbook = xlsxwriter.Workbook('report.xlsx')
        worksheet = workbook.add_worksheet('ЗВІТ 14ф')

        worksheet.write('J1', 'ЗВІТ ПО НОЗОЛОГІЯХ В РОЗРІЗІ РАЙОНІВ')
        worksheet.write('A2', 'МСЕК')
        worksheet.write('B2', str(self.DICT_KOMIS[num_msek]))
        worksheet.write('A3', 'ПЕРІОД')
        worksheet.write('B3', '{} - {}'.format(date_Z, date_PO))
        worksheet.write('A4', 'ШИФР')
        worksheet.write('B4', '{} - {}'.format(diagnoz_Z, diagnoz_PO))

        worksheet.write('A6', 'ТЕРИТОРІЯ')
        worksheet.write('B6', 'ВСЬОГО')
        worksheet.write('E5', 'ГРУПА')
        worksheet.write('D6', 'Iгр.')
        worksheet.write('C7', 'всього')
        worksheet.write('D7', 'Іa')
        worksheet.write('E7', 'Iб')
        worksheet.write('F6', 'ІІгр.')
        worksheet.write('G6', 'ІІІгр.')
        worksheet.write('H6', 'До 39р.')
        worksheet.write('I5', 'Від 40')
        worksheet.write('I6', 'до 55/60')
        
        worksheet.write('L5', 'Працюючі')
        worksheet.write('J6', 'ВСЬОГО')
        worksheet.write('L6', 'Iгр.')
        worksheet.write('K7', 'всього')
        worksheet.write('L7', 'Іa')
        worksheet.write('M7', 'Iб')
        worksheet.write('N6', 'ІІгр.')
        worksheet.write('O6', 'ІІІгр.')

        worksheet.write('P5', 'всьго')
        worksheet.write('P6', 'гр.6+')
        worksheet.write('P7', 'гр.7')
        worksheet.write('R5', 'Працездатний вік')
        worksheet.write('R6', 'Iгр.')
        worksheet.write('Q7', 'всього')
        worksheet.write('R7', 'Іa')
        worksheet.write('S7', 'Iб')
        worksheet.write('T6', 'ІІгр.')
        worksheet.write('U6', 'ІІІгр.')


        for key, value in first_part.items():
            worksheet.write(row, col, str(DICT_RAION[key]))
            worksheet.write(row, col + 1, int(''.join(str(value[0]))))
            worksheet.write(row, col + 2, int(''.join(str(value[1]))))
            worksheet.write(row, col + 3, int(''.join(str(value[2]))))
            worksheet.write(row, col + 4, int(''.join(str(value[3]))))
            worksheet.write(row, col + 5, int(''.join(str(value[4]))))
            worksheet.write(row, col + 6, int(''.join(str(value[5]))))
            worksheet.write(row, col + 7, int(''.join(str(value[6]))))
            worksheet.write(row, col + 8, int(''.join(str(value[7]))))
            worksheet.write(row, col + 9, int(''.join(str(value[8]))))
            worksheet.write(row, col + 10, int(''.join(str(value[9]))))
            worksheet.write(row, col + 11, int(''.join(str(value[10]))))
            worksheet.write(row, col + 12, int(''.join(str(value[11]))))
            worksheet.write(row, col + 13, int(''.join(str(value[12]))))
            worksheet.write(row, col + 14, int(''.join(str(value[13]))))
            worksheet.write(row, col + 15, int(''.join(str(value[14]))))
            worksheet.write(row, col + 16, int(''.join(str(value[15]))))
            worksheet.write(row, col + 17, int(''.join(str(value[16]))))
            worksheet.write(row, col + 18, int(''.join(str(value[17]))))
            worksheet.write(row, col + 19, int(''.join(str(value[18]))))
            worksheet.write(row, col + 20, int(''.join(str(value[19]))))
            row += 1
        
        for key, value in second_part.items():
            worksheet.write(row, col, str(DICT_CITY[key]))
            worksheet.write(row, col + 1, int(''.join(str(value[0]))))
            worksheet.write(row, col + 2, int(''.join(str(value[1]))))
            worksheet.write(row, col + 3, int(''.join(str(value[2]))))
            worksheet.write(row, col + 4, int(''.join(str(value[3]))))
            worksheet.write(row, col + 5, int(''.join(str(value[4]))))
            worksheet.write(row, col + 6, int(''.join(str(value[5]))))
            worksheet.write(row, col + 7, int(''.join(str(value[6]))))
            worksheet.write(row, col + 8, int(''.join(str(value[7]))))
            worksheet.write(row, col + 9, int(''.join(str(value[8]))))
            worksheet.write(row, col + 10, int(''.join(str(value[9]))))
            worksheet.write(row, col + 11, int(''.join(str(value[10]))))
            worksheet.write(row, col + 12, int(''.join(str(value[11]))))
            worksheet.write(row, col + 13, int(''.join(str(value[12]))))
            worksheet.write(row, col + 14, int(''.join(str(value[13]))))
            worksheet.write(row, col + 15, int(''.join(str(value[14]))))
            worksheet.write(row, col + 16, int(''.join(str(value[15]))))
            worksheet.write(row, col + 17, int(''.join(str(value[16]))))
            worksheet.write(row, col + 18, int(''.join(str(value[17]))))
            worksheet.write(row, col + 19, int(''.join(str(value[18]))))
            worksheet.write(row, col + 20, int(''.join(str(value[19]))))
            row += 1
            
            worksheet.write(row, col, 'ВСЬОГО')
            worksheet.write(row, col + 1, '=SUM(B1:B' + str(row) + ')')
            worksheet.write(row, col + 2, '=SUM(C1:C' + str(row) + ')')
            worksheet.write(row, col + 3, '=SUM(D1:D' + str(row) + ')')
            worksheet.write(row, col + 4, '=SUM(E1:E' + str(row) + ')')
            worksheet.write(row, col + 5, '=SUM(F1:F' + str(row) + ')')
            worksheet.write(row, col + 6, '=SUM(G1:G' + str(row) + ')')
            worksheet.write(row, col + 7, '=SUM(H1:H' + str(row) + ')')
            worksheet.write(row, col + 8, '=SUM(I1:I' + str(row) + ')')
            worksheet.write(row, col + 9, '=SUM(J1:J' + str(row) + ')')
            worksheet.write(row, col + 10, '=SUM(K1:K' + str(row) + ')')
            worksheet.write(row, col + 11, '=SUM(L1:L' + str(row) + ')')
            worksheet.write(row, col + 12, '=SUM(M1:M' + str(row) + ')')
            worksheet.write(row, col + 13, '=SUM(N1:N' + str(row) + ')')
            worksheet.write(row, col + 14, '=SUM(O1:O' + str(row) + ')')
            worksheet.write(row, col + 15, '=SUM(P1:P' + str(row) + ')')
            worksheet.write(row, col + 16, '=SUM(Q1:Q' + str(row) + ')')
            worksheet.write(row, col + 17, '=SUM(R1:R' + str(row) + ')')
            worksheet.write(row, col + 18, '=SUM(S1:S' + str(row) + ')')
            worksheet.write(row, col + 19, '=SUM(T1:T' + str(row) + ')')
            worksheet.write(row, col + 20, '=SUM(U1:U' + str(row) + ')')


        workbook.close()


if __name__ == "__main__":

    db = DB()