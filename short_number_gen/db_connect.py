import sqlite3,config,os,re
def init_db():
    if os.path.exists(config.db_file):
        c_sl3=sqlite3.connect(config.db_file)
        s_sl3 = c_sl3.cursor()
        return c_sl3
    else:
        c_sl3=sqlite3.connect(config.db_file)
        s_sl3 = c_sl3.cursor()
        s_sl3.execute('CREATE TABLE "tel" ("Name"  TEXT NOT NULL,"Type"  TEXT NOT NULL,"DH"  TEXT NOT NULL);')
        return c_sl3
def sqliteEscape(keyWord):
    keyWord = keyWord.replace("'", "''") 
    return keyWord
def add_item(cur,i):
    sql="INSERT INTO 'tel' VALUES ('{0}', '{1}', '{2}', '{3}');"
    sql=sql.format(i['XM'],i['BM'],i['DH'])
    s_sl3 = cur.cursor()
    s_sl3.execute(sql)
    cur.commit()
def find_item(cur,id,dh,bm):
    sql="SELECT count(*) FROM 'tel' where Name='{0}' and DH='{1}' and Type='{2}' ;"
    sql=sql.format(id,dh,bm)
    s_sl3 = cur.cursor()
    cussor=s_sl3.execute(sql)
    for r in cussor:
        return r[0]
    return -1


