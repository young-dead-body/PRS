import pyodbc
from openpyxl import load_workbook

def readExcelWriteLinter():
    book = load_workbook(
        filename="F:/JOB/PYTHON_NACHALO/Pars/PRS/RESOURSE/EXCEL/Мой средний балл.xlsx")  # name file к которому мы обращаемся
    sheetName = book['Лист1']  # имя листа который необходимо обработать
    for i in range(1, 45):
        if sheetName['B' + str(i)].value != None:
            cursor.execute(f"INSERT INTO PREDMET (ID_NAME, NAME) VALUES ({i}, '{sheetName['B' + str(i)].value}')")
            print(f"INSERT INTO PREDMET (ID_NAME, NAME) VALUES ({i}, '{sheetName['B' + str(i)].value}')")


conn = pyodbc.connect("""DRIVER=Linter 6.0 Driver;
                       SERVER=STUDENT;
                       DATABASE=STUDENT;
                       UID=ADMIN;
                       PWD=ADMIN""") # не хватает проверки

cursor = conn.cursor()

readExcelWriteLinter()

#cursor.execute("SELECT SURNAME FROM TEACHER")
#row = cursor.fetchone()
#while row:
#    print(row[0])
#    row = cursor.fetchone()