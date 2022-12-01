from tkinter import *
from tkinter import ttk
import pymysql
import pandas
from sqlalchemy import create_engine
import csv
import openpyxl


# --------------- window: --------------
root = Tk()
root.title("programme de gestion")
root.geometry("1350x690+1+1")
root.config(background="silver")
# root.iconbitmap("C:\\Users\\Dell2021\\Documents\\PythonScripts\\proj1\\etudiant.ico")
tiltle = Label(root, text="programme de gestion des etudiants", bg="#1AAECB",
               fg="white", font=("monoscope", 14))
tiltle.pack(fill=X)


# --------------- variables: --------------
id_var = StringVar()
name_var = StringVar()
mail_var = StringVar()
tel_var = StringVar()
gn_var = StringVar()
adr_var = StringVar()
rm_var = StringVar()
src_var = StringVar()
src_by = StringVar()

csv_data = []

# --------------- csv_file(): --------------


def csv_file():
    my_conn = create_engine("mysql+pymysql://root:""@localhost/student")
    sql = "select * from student"
    my_data = pandas.read_sql(sql, my_conn)
    my_data.to_csv(
        r"C:\Users\Dell2021\Documents\PythonScripts\proj1\students.csv", index=False)
    with open(r"C:\Users\Dell2021\Documents\PythonScripts\proj1\students.csv") as file_obj:
        reader = csv.reader(file_obj)
        for row in reader:
            csv_data.append(row)

    wb = openpyxl.Workbook()
    sheet = wb.active
    for row in csv_data:
        sheet.append(row)
    wb.save("student.xlsx")


# --------------- add_user(): --------------


def add_user():
    cnx = pymysql.connect(host="localhost", user="root",
                          password="", database="student")
    cur = cnx.cursor()
    cur.execute(
        "insert into student values (%s,%s,%s,%s,%s,%s)", (id_var.get(), name_var.get(), mail_var.get(), tel_var.get(), gn_var.get(), adr_var.get()))
    cnx.commit()
    fetch()
    clear()
    cnx.close()

# --------------- fetch(): --------------


def fetch():
    cnx = pymysql.connect(host="localhost", user="root",
                          password="", database="student")
    cur = cnx.cursor()
    cur.execute("select * from student")
    rows = cur.fetchall()
    if len(rows) != 0:
        trv.delete(*trv.get_children())
        for row in rows:
            trv.insert("", END, value=row)
        cnx.commit()
    cnx.close()


# --------------- delete(): --------------


def delete_user():
    cnx = pymysql.connect(host="localhost", user="root",
                          password="", database="student")
    cur = cnx.cursor()
    cur.execute("delete from student where id=%s", rm_var.get())
    cnx.commit()
    fetch()
    cnx.close()

# --------------- clear(): --------------


def clear():
    id_var.set("")
    name_var.set("")
    mail_var.set("")
    tel_var.set("")
    gn_var.set("")
    adr_var.set("")
    rm_var.set("")
    src_by.set("")
    src_var.set("")


# --------------- cursor(): --------------
def cursor(ev):
    cursor_row = trv.focus()
    contents = trv.item(cursor_row)
    row = contents["values"]
    id_var.set(row[0])
    name_var.set(row[1])
    mail_var.set(row[2])
    tel_var.set(row[3])
    gn_var.set(row[4])
    adr_var.set(row[5])


# --------------- update_user(): --------------
def update_user():
    cnx = pymysql.connect(host="localhost", user="root",
                          password="", database="student")
    cur = cnx.cursor()
    cur.execute("update student set name=%s, mail=%s, telephone=%s, gender=%s, addresse=%s  where id = %s",
                (name_var.get(), mail_var.get(), tel_var.get(), gn_var.get(), adr_var.get(), id_var.get()))
    cnx.commit()
    fetch()
    clear()
    cnx.close()

# --------------- search_user(): --------------


def search_user():
    cnx = pymysql.connect(host="localhost", user="root",
                          password="", database="student")
    cur = cnx.cursor()
    cur.execute("select * from student where "+str(src_by.get()) +
                " like '%"+str(src_var.get())+"%'")
    rows = cur.fetchall()
    if len(rows) != 0:
        trv.delete(*trv.get_children())
        for row in rows:
            trv.insert("", END, value=row)
        cnx.commit()
    cnx.close()


# --------------- manage_frame: --------------
mange_frame = Frame(root, bg="white")
mange_frame.place(x=1069, y=28, height=330, width=210)
# lbl_ch = Label(mange_frame, text="les champs", bg="#1AAECB", fg="white", font=("monoscope", 11))
# lbl_ch.pack(fill=X)
lbl_id = Label(mange_frame, text="id", bg="white")
lbl_id.pack()
ent_id = Entry(mange_frame, bd="2", justify="center", textvariable=id_var)
ent_id.pack()
lbl_name = Label(mange_frame, text="name", bg="white")
lbl_name.pack()
ent_name = Entry(mange_frame, bd="2", justify="center", textvariable=name_var)
ent_name.pack()
lbl_mail = Label(mange_frame, text="mail", bg="white")
lbl_mail.pack()
ent_mail = Entry(mange_frame, bd="2", justify="center", textvariable=mail_var)
ent_mail.pack()
lbl_tel = Label(mange_frame, text="telephone", bg="white")
lbl_tel.pack()
ent_tel = Entry(mange_frame, bd="2", justify="center", textvariable=tel_var)
ent_tel.pack()
lbl_gender = Label(mange_frame, text="gender", bg="white")
lbl_gender.pack()
combo_gender = ttk.Combobox(mange_frame, textvariable=gn_var)
combo_gender["value"] = ("male", "female")
combo_gender.pack()
lbl_adr = Label(mange_frame, text="addresse", bg="white")
lbl_adr.pack()
ent_adr = Entry(mange_frame, bd="2", justify="center", textvariable=adr_var)
ent_adr.pack()
lbl_rm = Label(mange_frame, text="remove someone by id", bg="white")
lbl_rm.pack()
ent_rm = Entry(mange_frame, bd="2", justify="center", textvariable=rm_var)
ent_rm.pack()

# --------------- button_frame: --------------
btn_frame = Frame(root, bg="white")
btn_frame.place(x=1069, y=358, height=400, width=210)
lbl = Label(btn_frame, text="les buttons", bg="#1AAECB")
lbl.pack(fill=X)

add_btn = Button(btn_frame, text="add user", bg="#AED6F1",
                 fg="white", command=add_user)
add_btn.place(x=33, y=50, width=150, height=30)

del_btn = Button(btn_frame, text="delete user",
                 bg="#AED6F1", fg="white", command=delete_user)
del_btn.place(x=33, y=83, height=30, width=150)

update_btn = Button(btn_frame, text="update user",
                    bg="#AED6F1", fg="white", command=update_user)
update_btn.place(x=33, y=116, height=30, width=150)

clr_btn = Button(btn_frame, text="clear(reset)",
                 bg="#AED6F1", fg="white", command=clear)
clr_btn.place(x=33, y=149, height=30, width=150)

about_btn = Button(btn_frame, text="refresh", bg="#AED6F1",
                   fg="white", command=fetch)
about_btn.place(x=33, y=182, height=30, width=150)

csv_btn = Button(btn_frame, text="csv file", bg="#AED6F1",
                 fg="white", command=csv_file)
csv_btn.place(x=33, y=215, height=30, width=150)

exit_btn = Button(btn_frame, text="exit", bg="#AED6F1",
                  fg="white", command=root.quit)
exit_btn.place(x=33, y=248, height=30, width=150)

# --------------- search_frame: --------------
src_frame = Frame(root, bg="white")
src_frame.place(x=1, y=28, height=50, width=1068)

lbl_name_ent = Label(src_frame, text="search for a user by :", bg="white")
lbl_name_ent.place(x=10, y=12)
combo_src = ttk.Combobox(src_frame, textvariable=src_by)
combo_src["value"] = ("id", "name", "mail", "telephone", "addresse")
combo_src.place(x=125, y=12, width=150)

ent_combo = Entry(src_frame, bd="2", justify="center", textvariable=src_var)
ent_combo.place(x=300, y=12, height=23, width=150)

btn_combo = Button(src_frame, text="search", bg="#AED6F1",
                   fg="white", command=search_user)
btn_combo.place(x=475, y=12, height=22, width=150)

# --------------- Data_frame: --------------
data_frame = Frame(root, bg="#ECF0F1")  # #ECF0F1
data_frame.place(x=0, y=78, height=578, width=1068)

scrl_x = Scrollbar(data_frame, orient=HORIZONTAL)
scrl_y = Scrollbar(data_frame, orient=VERTICAL)

trv = ttk.Treeview(data_frame,
                   columns=("id", "name", "mail", "tel", "gender", "addresse"),
                   xscrollcommand=scrl_x.set,
                   yscrollcommand=scrl_y.set)

trv.place(x=18, y=1, height=560, width=1060)
scrl_x.pack(side=BOTTOM, fill=X)
scrl_y.pack(side=LEFT, fill=Y)
scrl_x.config(command=trv.xview)
scrl_y.config(command=trv.yview)
# ----headings(table or treeview):
trv["show"] = "headings"
trv.heading("id", text="id")
trv.heading("name", text="name")
trv.heading("mail", text="mail")
trv.heading("tel", text="telephone")
trv.heading("gender", text="gender")
trv.heading("addresse", text="addresse")
trv.bind("<ButtonRelease-1>", cursor)  # ButtonRelease-1> ma3labalish wash dir
# -----width of headings:
trv.column("id", width=152)
trv.column("name", width=178)
trv.column("mail", width=178)
trv.column("tel", width=178)
trv.column("gender", width=178)
trv.column("addresse", width=178)

# --------------- refresh: --------------
fetch()

root.mainloop()


# pyinstaller.exe --onefile --windowed -i icone_name.ico python_file_name.py
