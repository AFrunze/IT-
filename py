import pypyodbc

connection = pypyodbc.connect("Driver={SQL Server}; Server=localhost;Port=1433;Database=ImportFromExcel")

cursor = connection.cursor()
#a
MySQLQuery1 = (""""
                SELECT *
                FROM dbo.base_dq
                where date_str > "01/07/2010"
               """)

cursor.execute(MySQLQuery1)
results1 = cursor.fetchall()
#print(results1)
#b
MySQLQuery2 = (""""
                SELECT *
                FROM dbo.base_dq
                where ag_name not like 'Интерфакс' and ag_name not like 'АМБест Компани'
               """)

cursor.execute(MySQLQuery2)
results2 = cursor.fetchall()
#print(results2)
#c
MySQLQuery3 = (""""
                SELECT *
                FROM dbo.base_dq
                where ent_name like '%банк%' or ent_name like '%bank%'
               """)

cursor.execute(MySQLQuery3)
results3 = cursor.fetchall()
#print(results3)
#d
MySQLQuery4 = (""""
                SELECT *
                FROM dbo.base_dq
                where ent_name not like '%банк%' and ent_name not like '%bank%'
               """)

cursor.execute(MySQLQuery4)
results4 = cursor.fetchall()
#print(results4)
#e
MySQLQuery5 = (""""
                SELECT *
                FROM dbo.base_dq
                where scale_typer like 'Nsc' or scale_typer like ''
               """)

cursor.execute(MySQLQuery5)
results5 = cursor.fetchall()
#print(results5)













connection.close()

