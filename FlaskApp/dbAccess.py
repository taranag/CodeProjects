import sqlite3

connection = sqlite3.connect('database.db')

cur = connection.cursor()

cur.execute('''select * from reports''')
print(cur.fetchall())

connection.commit()
connection.close()
