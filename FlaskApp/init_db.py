import sqlite3

connection = sqlite3.connect('database.db')


with open('schema.sql') as f:
    connection.executescript(f.read())

cur = connection.cursor()

cur.execute("INSERT INTO posts (title, content) VALUES (?, ?)",
            ('First Post', 'Content for the first post')
            )

cur.execute("INSERT INTO posts (title, content) VALUES (?, ?)",
            ('Second Post', 'Content for the second post')
            )

cur.execute("INSERT INTO reports (companyID, options, groupBy, startDate, endDate) VALUES (?, ?, ?, ?, ?)",
            ('51', 'Title, Download', 'level', '2022-06-01', '2022-06-15')
            )

connection.commit()
connection.close()
