# Cas09Vezbanje

import psycopg2 as psycopg2
import openpyxl as openpyxl
from tkinter import *
con=psycopg2.connect(database='LJUDI',
    user='postgres',
    password='itoip',
    host='localhost',
    port='5432'
    )

root= Tk()
root.geometry=('100x400')

b1=Button(master=root,text='Izvezi muskarce',
          width=15,
          height=1,
          command=lambda:muskarci())
b1.place(relx=0.1,rely=0.2)
b2=Button(master=root,text='Izvezi zene',
          width=15,
          height=1,
          command=lambda:zene())
b2.place(relx=0.1,rely=0.4)
b3=Button(master=root,text='Izbor',
          width=15,
          height=1,
          command=lambda:izbor())
b3.place(relx=0.1,rely=0.6)


def muskarci():
    cursor=con.cursor()
    cursor.execute("SELECT * FROM COVEK WHERE POL='MUSKO'")
    result=cursor.fetchall()
    cursor.close()
    con.close()
    print(result)
    wb=openpyxl.Workbook()
    ws=wb.active
    ws['A1'].value='JMBG'
    ws['B1'].value='Ime'
    ws['C1'].value='Prezime'
    ws['D1'].value='Broj godina'
    ws['E1'].value='pol'
    i=2
    for x in result:
        ws.cell(column=1,row=i).value=x[0]
        ws.cell(column=2,row=i).value=x[1]
        ws.cell(column=3,row=i).value=x[2]
        ws.cell(column=4,row=i).value=x[3]
        ws.cell(column=5,row=i).value=x[4]
        i=i+1
    wb.save(filename='muskarci.xlsx')



def zene():
    cursor=con.cursor()
    cursor.execute("SELECT * FROM COVEK WHERE POL='ZENSKO'")
    result=cursor.fetchall()
    cursor.close()
    con.close()
    print(result)

    wb=openpyxl.Workbook()
    ws=wb.active
    ws['A1'].value='JMBG'
    ws['B1'].value='Ime'
    ws['C1'].value='Prezime'
    ws['D1'].value='Broj godina'
    ws['E1'].value='pol'
    i=2
    for x in result:
        ws.cell(column=1,row=i).value=x[0]
        ws.cell(column=2,row=i).value=x[1]
        ws.cell(column=3,row=i).value=x[2]
        ws.cell(column=4,row=i).value=x[3]
        ws.cell(column=5,row=i).value=x[4]
        i=i+1
    wb.save(filename='zene.xlsx')

def dodajclana():
    IDc=input('Unesite JMBG:  ')
    imec=input('Unesite Ime:  ')
    prezc=input('Unesite Prezime:  ')
    godinec=input('Unesite godine:  ')
    polc=input('Unesite pol:  ')
    cursor=con.cursor()
    cursor.execute('''INSERT INTO COVEK (JMBG, IME,PREZIME,BROJ_GODINA,POL) VALUES ('{}','{}','{}',{},'{}');'''.format(IDc,imec,prezc,eval(godinec),polc))
    con.commit()
    cursor.close()
    con.close()


def izbor():
    krit=input("Unesi kriterijum: ")
    cursor=con.cursor()
    cursor.execute("SELECT * FROM COVEK ORDER BY {}".format(krit))
    result=cursor.fetchall()
    cursor.close()
    con.close()
    print(result)
    wb=openpyxl.Workbook()
    ws=wb.active
    ws['A1'].value='JMBG'
    ws['B1'].value='Ime'
    ws['C1'].value='Prezime'
    ws['D1'].value='Broj godina'
    ws['E1'].value='pol'
    i=2
    for x in result:
        ws.cell(column=1,row=i).value=x[0]
        ws.cell(column=2,row=i).value=x[1]
        ws.cell(column=3,row=i).value=x[2]
        ws.cell(column=4,row=i).value=x[3]
        ws.cell(column=5,row=i).value=x[4]
        i=i+1
    wb.save(filename='redosled.xlsx')


menubar=Menu(master=root)
accountmenu=Menu(menubar,tearoff=0)

accountmenu.add_command(label="Dodaj",command=lambda:dodajclana())
menubar.add_cascade(label="Admin", menu=accountmenu)
root.config(menu=menubar)


root.mainloop()
