import sqlite3,config,os,re
def init_db():
    if os.path.exists(config.db_file):
        c_sl3=sqlite3.connect(config.db_file)
        s_sl3 = c_sl3.cursor()
        return c_sl3
    else:
        c_sl3=sqlite3.connect(config.db_file)
        s_sl3 = c_sl3.cursor()
        s_sl3.execute('CREATE TABLE "news" ("news_ID"  TEXT NOT NULL,"news_title"  TEXT NOT NULL,"news_LMID"  TEXT NOT NULL,"news_LMNAME"  TEXT NOT NULL,"news_FBBM"  TEXT NOT NULL,"news_DATE"  date NOT NULL,"news_content"  TEXT,"news_url"  TEXT,"news_FJ"  TEXT);')
        return c_sl3
def sqliteEscape(keyWord):
    keyWord = keyWord.replace("'", "''") 
    return keyWord
def add_item(cur,i):
    sql="INSERT INTO 'news' VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}');"
    sql=sql.format(i['news_ID'],i['news_title'],i['news_LMID'],i['news_LMNAME'],i['news_FBBM'],i['news_DATE'],sqliteEscape(i['news_content']),sqliteEscape(i['news_url']),i['news_FJ'])
    s_sl3 = cur.cursor()
    s_sl3.execute(sql)
    cur.commit()
def find_item(cur,id):
    sql="SELECT count(*) FROM 'news' where news_ID='{}' ;"
    sql=sql.format(id)
    s_sl3 = cur.cursor()
    cussor=s_sl3.execute(sql)
    for r in cussor:
        return r[0]
    return -1


