from os import getcwd, listdir
import sqlite3


num_msek = 1
date_Z = '2018-01-01'
date_PO = '2018-12-31'
KODRAION = 7
diagnoz_Z = 'A00.0'
diagnoz_PO = 'S00.0'

path = getcwd() + '/db_files/'
conn = sqlite3.connect(path + 'db_{}.db'.format(num_msek))
cur = conn.cursor()


XZ = cur.execute("""SELECT COUNT(roglad) FROM db WHERE oglad = 1 and datezakl BETWEEN ? and ? and kodraion = ? and roglad BETWEEN 0 and 3 and diagnoz BETWEEN ? and ?""", [date_Z, date_PO, KODRAION, diagnoz_Z, diagnoz_PO]).fetchone()

print(XZ[0])