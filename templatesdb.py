import sqlite3 as sq


def createdb():
    with sq.connect("templates.db") as con:
        cur = con.cursor()
        cur.executescript("""
        CREATE TABLE IF NOT EXISTS templates (
            templ_num INTEGER DEFAULT 0,
            description TEXT NOT NULL,
            paragraphs TEXT NOT NULL);
        CREATE TRIGGER IF NOT EXISTS after_insert AFTER INSERT ON templates
            BEGIN
            UPDATE templates SET templ_num = IFNULL((SELECT MAX(templ_num) FROM templates), 0) + 1 WHERE rowid = NEW.ROWID;
            END;
        CREATE TRIGGER IF NOT EXISTS after_delete AFTER DELETE ON templates
            BEGIN
            UPDATE templates SET templ_num = templ_num - 1 WHERE templ_num > OLD.templ_num;
            END;
        """)


def inserttemplateindb(tpl):
    with sq.connect("templates.db") as con:
        cur = con.cursor()
        cur.execute("INSERT INTO templates (description, paragraphs) VALUES (?, ?)", tpl)


def gettemplateslistfromdb():
    templatelistwithdict = []
    with sq.connect("templates.db") as con:
        cur = con.cursor()
        cur.execute("SELECT * FROM templates")
        for row in cur:
            templatelistwithdict.append({'templnum': row[0], 'desctempl': row[1]})
    return templatelistwithdict


def deletetemplatefromdb(num):
    with sq.connect("templates.db") as con:
        cur = con.cursor()
        cur.execute(f"DELETE FROM templates where templ_num = {num}")


def gettemplatetextfromdb(num):
    with sq.connect("templates.db") as con:
        cur = con.cursor()
        cur.execute(f"SELECT paragraphs FROM templates where templ_num = {num}")
        text = cur.fetchone()
        return text[0]


def gettemplatedescfromdb(num):
    with sq.connect("templates.db") as con:
        cur = con.cursor()
        cur.execute(f"SELECT description FROM templates where templ_num = {num}")
        text = cur.fetchone()
        return text[0]


def setupdatetemplateindb(num, desc, text):
    with sq.connect("templates.db") as con:
        cur = con.cursor()
        cur.execute(f"UPDATE templates SET description = '{desc}', paragraphs = '{text}' where templ_num = {num}")
        