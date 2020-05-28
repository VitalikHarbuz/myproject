import sqlite3

"""[(0, 'kod', 'TEXT', 0, None, 0), (1, 'oglad', 'INTEGER', 0, None, 0), (2, 'datezakl', 'TEXT', 0, None, 0), (3, 'famil', 'TEXT', 0, None, 0), (4, 'name', 'TEXT', 0, None, 0), 
(5, 'tato', 'TEXT', 0, None, 0), (6, 'pol', 'INTEGER', 0, None, 0), (7, 'rod', 'TEXT', 0, None, 0), (8, 'kodsity', 'INTEGER', 0, None, 0), (9, 'kodraion', 'INTEGER', 0, None, 0), 
(10, 'selo', 'TEXT', 0, None, 0), (11, 'street', 'TEXT', 0, None, 0), (12, 'woker', 'INTEGER', 0, None, 0), (13, 'social', 'INTEGER', 0, None, 0), (14, 'special', 'TEXT', 0, None, 0), 
(15, 'mesto', 'TEXT', 0, None, 0), (16, 'ministr', 'TEXT', 0, None, 0), (17, 'organ', 'TEXT', 0, None, 0), (18, 'lpu', 'TEXT', 0, None, 0), (19, 'moglad', 'INTEGER', 0, None, 0), 
(20, 'grupinv', 'INTEGER', 0, None, 0), (21, 'toglad', 'INTEGER', 0, None, 0), (22, 'roglad', 'INTEGER', 0, None, 0), (23, 'invalid', 'INTEGER', 0, None, 0), 
(24, 'poteri', 'INTEGER', 0, None, 0), (25, 'prichin', 'INTEGER', 0, None, 0), (26, 'faktor', 'TEXT', 0, None, 0), (27, 'diagnoz', 'TEXT', 0, None, 0), 
(28, 'expert', 'INTEGER', 0, None, 0), (29, 'lik', 'INTEGER', 0, None, 0), (30, 'rebmed', 'TEXT', 0, None, 0), (31, 'rebwork', 'TEXT', 0, None, 
0), (32, 'rebhelp', 'TEXT', 0, None, 0), (33, 'prog1', 'TEXT', 0, None, 0), (34, 'prog2', 'TEXT', 0, None, 0), (35, 'pr', 'INTEGER', 0, None, 0), 
(36, 'prog3', 'TEXT', 0, None, 0), (37, 'poterist', 'TEXT', 0, None, 0), (38, 'prre', 'TEXT', 0, None, 0), (39, 'rogladi', 'INTEGER', 0, None, 0), 
(40, 'tran', 'TEXT', 0, None, 0), (41, 'grupinvi', 'TEXT', 0, None, 0)]"""
#cur.execute("""PRAGMA table_info(db)""") #переглянути список стовпців в таблиці
conn = sqlite3.connect('db_1.db')
cur = conn.cursor()

'''ins = cur.execute("""INSERT INTO TABLE db(kod, oglad, datezakl, famil, name, tato, pol, rod, kodsity, kodraion, 
                                          selo, street, woker, social, special, mesto, ministr, organ, lpu, moglad, 
                                          grupinv, toglad, roglad, invalid, poteri, prichin, faktor, diagnoz, expert, lik, 
                                          rebmed, rebwork, rebhelp, prog1, prog2, pr, prog3, poterist, prre, rogladi, 
                                          tran, grupinvi) VALUES (?,?,?,?,?,?,?,?,?,?,
                                                                  ?,?,?,?,?,?,?,?,?,?,
                                                                  ?,?,?,?,?,?,?,?,?,?,
                                                                  ?,?,?,?,?,?,?,?,?,?,
                                                                  ?,?)""")

create = cur.execute("""CREATE TABLE IF NOT EXISTS db(kod, oglad, datezakl, famil, name, tato, pol, rod, kodsity, kodraion, 
												selo, street, woker, social, special, mesto, ministr, organ, lpu,
												moglad, grupinv, grupinvi, toglad, roglad, rogladi, invalid, poteri, prichin, faktor, diagnoz, expert,
												rebwork, rebhelp, prog2, prog3)""")
conn.commit()
cur.close()
conn.close()'''

f_name = "КУШНІР"
name = "ВАСИЛЬ"
surname = "ВАСИЛЬОВИЧ"

conn = sqlite3.connect('db_1.db')
cur = conn.cursor()
exe = cur.execute("""SELECT * FROM db""")
[print(row, '\n') for row in exe.fetchall()]
#ins = cur.execute("""INSERT INTO db(famil, name, tato) VALUES(?, ?, ?)""", [f_name, name, surname])
<<<<<<< HEAD
#conn.commit()
=======
#conn.commit()
>>>>>>> 02f5a6e3ac1c8f33e54b2d036420fd64fa89e35f
