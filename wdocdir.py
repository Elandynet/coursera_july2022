import mysql.connector
from docx2python import docx2python
import os
import csv

cnx = mysql.connector.connect(user = 'root',
    password = 'aL589228',
    host = 'localhost',
    database = 'glossary')
cur = cnx.cursor()
#borrar = {',':'', '\:':'', '\'':'', '\"':'', '(':'', ')':'', '!':'', '?':'', '*':'', '.':'', '...':'', '“':'', '”':'', '[':'', ']':'' }
borrar = (',', ':',  ';', '(', ')', '!', '?', '*', '.', '...', '“', '”', '[', ']')
os.chdir("C:\Webdeb\Apache24\htdocs\python\words\docs")
filenames = os.listdir()
newords = []
for value in filenames:
    count = 0
    count2 = 0
    name = value
    doc = docx2python(value)
    texto = doc.text
    for x in borrar:
        texto = texto.replace(x, "")
    texto = texto.lower()
    words = texto.split()
    #print(words)
    #words = list(set(words)) 
    status = "Processing " + name
    print(status)
    for value in words:
        if len(value) < 23:
             count2 = count2 + 1
             find = "SELECT * FROM words WHERE word = %s LIMIT 1"
             cur.execute(find, (value,))
             result  = cur.fetchall()
             if result:
                 new_count = result[0][2] + 1
                 update = "UPDATE words SET occurance = %s WHERE word_id = %s"
                 data = (new_count, result[0][0])
                 cur.execute(update, data)
                 cnx.commit()
             else:
                 insert = "INSERT INTO words (word, occurance) VALUES (%s, %s)"
                 data = (value, 1)
                 cur.execute(insert, data)
                 cnx.commit()
                 count = count + 1
                 neword = [value]
                 newords.append(neword)
    insert = "INSERT INTO documents (docname, totwords, newords) VALUES (%s, %s, %s)"
    data = (name, count2, count)
    cur.execute(insert, data)
    cnx.commit()
    status = "Completed\n" + str(len(words)) + " words in document\n" + str(count2) + " words processed\n" + str(count) + " words added\n"
    print(status)
#csvname = "100-102.csv"
#with open(csvname, 'w') as f:
#    write = csv.writer(f)
#   write.writerows(newords)
cur.close()
cnx.close()
