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

var=StringVar()
w1 = Radiobutton ( master=root,text='Ime',variable=var,value="Ime")
w1.place(relx=0.5,rely=0.2)

w2 = Radiobutton ( master=root,text='Prezime',variable=var,value="Prezime")
w2.place(relx=0.5,rely=0.5)

w3 = Radiobutton (master=root,text='Godina',variable=var,value="Godina")
w3.place(relx=0.5,rely=0.7)

b4=Button(master=root,text='Export',
          width=15,
          height=1,
          command=lambda:izbor(var.get()))
b4.place(relx=0.5,rely=0.8)

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
    dodaj=Toplevel(root)
    dodaj.geometry('200x2009')
    l11=Label(dodaj,text='JMBG')
    l11.place(relx=0.1,rely=0.1)
    l12=Label(dodaj,text='Ime')
    l12.place(relx=0.1,rely=0.2)
    l13=Label(dodaj,text='Prezime')
    l13.place(relx=0.1,rely=0.3)
    l14=Label(dodaj,text='Godine')
    l14.place(relx=0.1,rely=0.4)
    l15=Label(dodaj,text='Pol')
    l15.place(relx=0.1,rely=0.5)

    e11=Entry(dodaj)
    e11.place(relx=0.5,rely=0.1)
    e12=Entry(dodaj)
    e12.place(relx=0.5,rely=0.2)
    e13=Entry(dodaj)
    e13.place(relx=0.5,rely=0.3)
    e14=Entry(dodaj)
    e14.place(relx=0.5,rely=0.4)
    e15=Entry(dodaj)
    e15.place(relx=0.5,rely=0.5)
    b11=Button(dodaj,text='Dodaj',command=lambda:[aa=e11.get(),imec=e12.get(),prezc=e13.get(),godinec=eval(e14.get()),polc=e15.get()])
    b11.place(relx=0.5,rely=0.8)
    cursor=con.cursor()
    cursor.execute('''INSERT INTO COVEK (JMBG, IME,PREZIME,BROJ_GODINA,POL) VALUES ('{}','{}','{}',{},'{}');'''.format(aa,imec,prezc,eval(godinec),polc))
    con.commit()
    cursor.close()
    con.close()


def izbor(krit):
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
