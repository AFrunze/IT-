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

#создание словоря для  рейтингов, в эксельке 300 уникальных значений
#данные о рейтингах взяты из википедии 

d = {"A"="0","A-"="0","A(kaz)"="0","A(rus)"="0","A-(rus)"="0","A(ukr)"="0","A-(ukr)"="0","A.mfi"="0","A+"="0","A+(kaz)"="0","A+(rus)"="0","A+(ukr)"="0","A+.mfi"="0","A++"="0","A++.mfi"="0","A1"="0","A-1"="0","A1(cr)"="0","A1.ru"="0","A-1+"="0","kzA"="0",
"kzA-"="0",
"kzA+"="0",
"A2"="1",
"A-2"="1",
"A2(cr)"="1",
"A2.ru"="1",
"A21"="1",
"A21.ru"="1",
"A22"="1",
"A22.ru"="1",
"A23"="1",
"kzAA"="1",
"kzAA-"="1",
"kzAA+"="1",
"A23.ru"="2",
"a3"="2",
"A-3"="2",
"A3(cr)"="2",
"A3.ru"="2",
"A31"="2",
"A31.ru"="2",
"A32"="2",
"A32.ru"="2",
"A33"="2",
"A33.ru"="2",
"AA"="2",
"AA-"="2",
"AA(kaz)"="2",
"AA-(kaz)"="2",
"AA(rus)"="2",
"AA-(rus)"="2",
"AA(ukr)"="2",
"AA-(ukr)"="2",
"AA+"="2",
"AA+(kaz)"="2",
"AA+(rus)"="2",
"AA+(ukr)"="2",
"Aa1"="2",
"Aa1(cr)"="2",
"Aa1.ru"="2",
"Aa2"="2",
"Aa2(cr)"="2",
"Aa2.ru"="2",
"kzAAA"="2",
"Aa3"="3",
"Aa3(cr)"="3",
"Aa3.ru"="3",
"AAA"="3",
"Aaa(cr)"="3",
"AAA(kaz)"="3",
"AAA(rus)"="3",
"AAA(ukr)"="3",
"Aaa.ru"="3",
"iAA-"="3",
"b"="4",
"B-"="4",
"B(rus)"="4",
"B-(rus)"="4",
"B(ukr)"="4",
"B-(ukr)"="4",
"B/C"="4",
"B+"="4",
"B+(kaz)"="4",
"B+(rus)"="4",
"B+.mfi"="4",
"B++"="4",
"B++.mfi"="4",
"B1"="4",
"B-1"="4",
"B1(cr)"="4",
"B1.kz"="4",
"B1.ua"="4",
"B11"="4",
"B11.ru"="4",
"B12"="4",
"B12.ru"="4",
"B13"="4",
"B13.ru"="4",
"B1-PD"="4",
"kzB-"="4",
"B2"="5",
"B2(cr)"="5",
"B2.ru"="5",
"B2.ua"="5",
"B21"="5",
"B21.ru"="5",
"B22"="5",
"B22.ru"="5",
"B23"="5",
"B23.ru"="5",
"B2-PD"="5",
"kzBB"="5",
"kzBB-"="5",
"kzBB+"="5",
"kzBBB"="5",
"kzBBB+"="5",
"B3"="6",
"B3(cr)"="6",
"B3.kz"="6",
"B3.ru"="6",
"B3.ua"="6",
"B31"="6",
"B31.ru"="6",
"B32"="6",
"B32.ru"="6",
"B33"="6",
"B33.ru"="6",
"B3-PD"="6",
"iB+"="6",
"iBB"="6",
"iBB+"="6",
"iBBB"="6",
"iBBB-"="6",
"ba1"="7",
"Ba1(cr)"="7",
"Ba1.ru"="7",
"Ba1-PD"="7",
"Ba2"="7",
"Ba2(cr)"="7",
"Ba2.kz"="7",
"Ba2.ru"="7",
"Ba2-PD"="7",
"Ba3"="7",
"Ba3(cr)"="7",
"Ba3.kz"="7",
"Ba3.ru"="7",
"Ba3-PD"="7",
"Baa1"="8",
"Baa1(cr)"="8",
"Baa1.ru"="8",
"Baa2"="9",
"Baa2(cr)"="9",
"Baa2.ru"="9",
"Baa3"="10",
"Baa3(cr)"="10",
"Baa3.kz"="10",
"Baa3.ru"="10",
"Baa3-PD"="10",
"BB"="10",
"BB-"="10",
"BB(kaz)"="10",
"BB-(kaz)"="10",
"BB(rus)"="10",
"BB-(rus)"="10",
"BB(ukr)"="10",
"BB+"="10",
"BB+(kaz)"="10",
"BB+(rus)"="10",
"BB+(ukr)"="10",
"BBB"="11",
"BBB-"="11",
"BBB(rus)"="11",
"BBB-(rus)"="11",
"BBB(ukr)"="11",
"BBB-(ukr)"="11",
"BBB+"="11",
"BBB+(rus)"="11",
"BBB+(ukr)"="11",
"C"="12",
"C-"="12",
"C(cr)"="12",
"C(rus)"="12",
"C(ukr)"="12",
"C.ru"="12",
"C.ua"="12",
"C/D"="12",
"C+"="13",
"C++"="13",
"C++.mfi"="13",
"C11"="13",
"C11.ru"="13",
"C12"="13",
"C12.ru"="13",
"C13"="13",
"C13.ru"="13",
"C2"="13",
"C2.ru"="13",
"C3"="13",
"C3.ru"="13",
"ca"="13",
"Ca.ru"="13",
"Ca.ua"="13",
"Caa1"="14",
"Caa1(cr)"="14",
"Caa1.ru"="14",
"Caa1-PD"="14",
"Caa2"="14",
"Caa2(cr)"="14",
"Caa2.ru"="14",
"Caa2.ua"="14",
"Caa2-PD"="14",
"Caa3"="14",
"Caa3(cr)"="14",
"Caa3.ru"="14",
"Caa3.ua"="14",
"Caa3-PD"="14",
"Ca-PD"="14",
"CC"="14",
"CC(rus)"="14",
"CC+"="14",
"ccc"="14",
"CCC-"="14",
"CCC(rus)"="14",
"CCC-(rus)"="14",
"CCC(ukr)"="14",
"CCC+"="14",
"C-PD"="14",
"D"="15",
"D-"="15",
"D(rus)"="15",
"D(ukr)"="15",
"D.ru"="15",
"D/E"="15",
"D+"="15",
"D-PD"="15",
"E"="16",
"E+"="16",
"F"="16",
"F1"="16",
"F1(rus)"="16",
"F1+"="16",
"F1+(kaz)"="16",
"F1+(rus)"="16",
"F2"="16",
"F2(rus)"="16",
"F3"="16",
"kzD"="16",
"NoFloor"="16",
"Not Prime"="16",
"Not Prime(cr)"="16",
"NR"="16",
"Prime-1"="16",
"Prime-1(cr)"="16",
"Prime-2"="16",
"Prime-2(cr)"="16",
"Prime-3"="16",
"Prime-3(cr)"="16",
"R"="16",
"RD"="16",
"RD(rus)"="16",
"RD(ukr)"="16",
"Removed"="16",
"ruA"="16",
"ruA-"="16",
"ruA+"="16",
"ruAA"="16",
"ruAA-"="16",
"ruAA+"="16",
"ruAAA"="16",
"ruB"="16",
"ruB-"="16",
"ruB+"="16",
"ruBB"="16",
"ruBB-"="16",
"ruBB+"="16",
"ruBBB"="16",
"ruBBB-"="16",
"ruBBB+"="16",
"ruCC"="16",
"ruCCC"="16",
"ruCCC-"="16",
"ruCCC+"="16",
"ruD"="16",
"ruR"="16",
"ruSD"="16",
"SD"="16",
"uaB+"="16",
"uaBB-"="16",
"uaBB+"="16",
"uaBBB-"="16",
"uaCCC-"="16",
"uaCCC+"="16",
"GAMMA-5"="17",
"GAMMA-5+"="17",
"GAMMA-6+"="17",
"GAMMA-7+"="17",
"Приостановлен"="17",
"Снят"="17"}

#перекодировка grade в соответствии со словарем

MySQLQuery6 = (""
                SELECT grade
                FROM dbo.base_dq
               "")
               
cursor.execute(MySQLQuery6)
results6 = cursor.fetchall() 

for r in results6:
  if r[1] in d


connection.close()
