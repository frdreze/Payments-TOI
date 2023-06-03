from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from docxtpl import DocxTemplate
import csv

tab_dat_edit_flag = False
things = ["ИМЯ", "ИНН", "КПП","БАНК", "КОРРСЧЁТ", "БИК"]
payer = 0
recipient = 0

# Форматно-логический контроль, выдающий правильность
def logic_num(num, bik):
    bik_3 = str(bik)[-3:]
    num = list(str(num))
    ktr = int(num[8])
    num[8] = 0
    ves = "713"*6 +"71"
    bik = [int(i)*int(j) for i, j in zip(ves[:3], bik_3)] 
    bik = [str(i)[-1] for i in bik]
    sum_bik = sum(list(map(int, bik)))
    num = [int(i)*int(j) for i, j in zip(ves, num)]
    num = [str(i)[-1] for i in num]
    sum_num = sum(list(map(int, num)))
    res = (sum_num+sum_bik)%10*3%10
    return res==ktr
def logic_inn(inn):
    inn = str(inn)
    if len(inn)==10:
        vesa_inn = (2, 4, 10, 3, 5, 9, 4, 6, 8)
        return sum([int(i)*j for i,j in zip(inn[:-1], vesa_inn)])%11%10==inn[-1]
    if len(inn)==12:
        vesa_inn_1 = [7, 2, 4, 10, 3, 5, 9, 4, 6, 8]
        vesa_inn_2 = [3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8]
        step_1 = sum([int(i)*j for i,j in zip(inn[:-2], vesa_inn_1)])%11%10
        step_2 = sum([int(i)*j for i,j in zip(inn[:-1], vesa_inn_2)])%11%10
        return (int(inn[-2])==step_1) and (int(inn[-1])==step_2)




# ВВОД НОВЫХ ЗНАЧЕНИЙ ДЛЯ ПЛАТЕЛЬЩИКА И ПОЛУЧАТЕЛЯ
def level_subfunction(dictionary):
    global tab_dat_edit_flag
    tab_dat.update(dictionary)
    tab_dat_edit_flag = True
    messagebox.showinfo("Данные добавлены", "Перейдите обратно в главную форму")
    combox_update = list(tab_dat.keys())
    payer_combobox.configure(values=combox_update)
    recipient_combobox.configure(values=combox_update)
def lev_pay():
    level_payer = Toplevel()
    level_payer.title("Плательщик")
    level_payer_list = []
    for idex,i in enumerate(["Счёт"]+ things +["Пароль"]):
        en = Entry(level_payer)
        la = Label(level_payer, text=i)
        en.grid(column=1, row=idex)
        la.grid(column=0, row=idex)
        level_payer_list.append(en)
    button_pay_check = Button(level_payer, text="Проверка данных", command=lambda: level_subfunction({level_payer_list[0].get(): [i.get() for i in level_payer_list[1:-1]]}) if level_payer_list[-1].get()=="Serezha_not_a_boar" and ((len(level_payer_list[2].get())==12 and len(level_payer_list[3].get())==0) or (len(level_payer_list[2].get())==10 and len(level_payer_list[3].get())==9)) and logic_inn(level_payer_list[2].get()) and logic_num(level_payer_list[0].get(), level_payer_list[-2].get()) else messagebox.showerror("Ошибка ввода", "Проверьте правильность БИК, счёта, ИНН, пароля. Может для физлиц вы указали КПП"))
    button_pay_check.grid(row=8, column=0, columnspan=2)
    print(tab_dat)
    # return logic_inn(payer_new.get(), tab_dat[payer_new.get()][-1])
def lev_rec():
    level_recipient = Toplevel()
    level_recipient.title("Получатель")
    level_recipient_list = []
    for idex,i in enumerate(["Счёт"]+ things +["Пароль"]):
        en = Entry(level_recipient)
        la = Label(level_recipient, text=i)
        en.grid(column=1, row=idex)
        la.grid(column=0, row=idex)
        level_recipient_list.append(en)
    button_recipe_check = Button(level_recipient, text="Проверка данных", command=lambda: level_subfunction({level_recipient_list[0].get(): [i.get() for i in level_recipient_list[1:-1]]}) if level_recipient_list[-1].get()=="Serezha_not_a_boar" and ((len(level_recipient_list[2].get())==12 and len(level_recipient_list[3].get())==0) or (len(level_recipient_list[2].get())==10 and len(level_recipient_list[3].get())==9)) and logic_inn(level_recipient_list[2].get()) and logic_num(level_recipient_list[0].get(), level_recipient_list[-2].get()) else messagebox.showerror("Ошибка ввода", "Проверьте правильность БИК, счёта, ИНН, пароля. Может для физлиц вы указали КПП"))
    button_recipe_check.grid(row=8, column=0, columnspan=2)



# Считываем словарь из системы
tab_dat = {}
header = ""
context_fill = {}
def table_path_func():
    global tab_dat, header
    filetypes = (
        ('csv files', '*.csv'),
        ('All files', '*.*')
    )
    table_path = filedialog.askopenfile(mode="r", filetypes=filetypes)
    table_file = csv.reader(table_path)
    header = next(table_file)
    for row in table_file:
        tab_dat[row[0]] = row[1:]
    combo = [i for i in tab_dat.keys()] + []
    payer_combobox["values"] = combo
    # payer_combobox.current(0)
    recipient_combobox["values"] = combo
    # recipient_combobox.current(0)
    print(tab_dat)



root = Tk()
root.geometry("1000x800")
root.title("Формирование платежки")
Frame_info = Frame(root,highlightthickness=2, highlightbackground="blue")
Frame_info.grid(column=0, row=0)

tab_dat_edit_flag = False
defaults = {}
summa_value = 0




# Создание таблицы
columns = ("Плат/получ?", "Имя", "ИНН", "КПП", "Банк", "Коррсчёт банка", "БИК банк")
tree = ttk.Treeview(root, columns=columns, show="headings")
for i in columns:
    tree.heading(i, text=i)
for i in ["ИНН", "КПП", "Банк"]: tree.column(i, width=150)
tree.grid(column=0, row=2)


def take_defaults():
    global defaults
    filetypes = (
        ('text files', '*.txt'),
        ('All files', '*.*')
    )
    defaults_file = filedialog.askopenfile(mode="r", filetypes=filetypes)
    defaults = {i[:i.find(":")]: i[i.find(":")+1:-1] for i in defaults_file.readlines()}
    context_fill.update(defaults)
def txt_report():
    txt_report_file = filedialog.asksaveasfile(mode="w", defaultextension=".txt")
    for i in context_fill.keys():
        txt_report_file.write(f"{i} --- {context_fill[i]}\n")
def pdf_report(): pass
def save_tab_dat():
    file = filedialog.asksaveasfile(defaultextension=".csv")
    file.write(",".join(header)+"\n")
    for i in tab_dat.keys():
        file.write(f"{i},{','.join(tab_dat[i])}\n")
        print(tab_dat)
def docx_report():
    docx_template = filedialog.askopenfilename(defaultextension=".docx")
    docx_template = DocxTemplate(docx_template)
    docx_template.render(context_fill)
    docx_template.render(context_fill)
    docx_report_file = filedialog.asksaveasfilename(defaultextension=".docx")
    docx_template.save(docx_report_file)


def good_ending():
    global context_fill
    ending = Toplevel()
    ending.title("Конец работы")
    list_payer = tab_dat[payer].copy()
    list_payer.insert(0, payer)
    list_recipient = tab_dat[recipient].copy()
    list_recipient.insert(0, recipient)
    context_fill.update({"Назначение_платежа": forwhat_entry.get()})
    dict_1 = {i:j for i,j in zip(["Плат_"+i for i in header], list_payer)}
    dict_2 = {i:j for i,j in zip(["Получ_" + i for i in header], list_recipient)}
    context_fill.update(dict_1)
    context_fill.update(dict_2)
    try:
        summa_value = float(sum_entry.get())
    except:
        messagebox.showerror("Проверьте значение суммы", "Попробуйте исправить сумму")
    context_fill.update({"Сумма_платежа": summa_value})

    txt_report_button = Button(ending, text="Экспортировать в txt", command=txt_report)
    txt_report_button.grid(row=0, column=1)
    # pdf_report_button = Button(ending, text="Экспорт в инфографическом виде", command=pdf_report)
    # pdf_report_button.grid()
    take_defaults_button = Button(ending, text="Экспортировать данные формы платежного поручения по умолчанию для формирования", command=take_defaults)
    take_defaults_button.grid(row=0, column=0)
    docs_report_button = Button(ending, text="Экспорт в формате docx", command=docx_report)
    docs_report_button.grid(row=1, column=0)
    print(tab_dat)
    if tab_dat_edit_flag==True:
        save_tab_dat_button = Button(ending, text="Сохранить изменения в базе плательщиков и получателей", command=save_tab_dat)
        save_tab_dat_button.grid(row=1, column=2)



    
def maketable():
    global payer, recipient
    payer = payer_combobox.get()
    recipient = recipient_combobox.get()
    if recipient!=payer:    
        for i in tree.get_children(): tree.delete(i)
        pays = ["Плательщик"]+tab_dat[payer]
        recps = ["Получатель"]+tab_dat[recipient]
        tree.insert("", 0, values=pays)
        tree.insert("", 0, values=recps)
    else: messagebox.showerror("Одинаковые счета", "Посмотрите на выпадающий список и исправьте ошибку")
    print(tab_dat)


Label(Frame_info, text="Номер счёта получателя").grid(column=0, row=2)
recipient_new = Button(Frame_info, text="Ввод нового получателя или редактирование старого", command=lev_rec)
recipient_new.grid(column=2, row=2)
recipient_combobox = ttk.Combobox(Frame_info, values=[])
recipient_combobox.grid(column=1, row=2)
Label(Frame_info, text="Назначение платежа").grid(column=1, row=0)
forwhat_entry = Entry(Frame_info)
forwhat_entry.grid(column=2, row=0)
Label(Frame_info, text="Номер счёта плательщика").grid(column=0, row=1)
Label(Frame_info, text="Введите сумму через точку").grid(row=3, column=0)
sum_entry = Entry(Frame_info)
sum_entry.grid(row=3, column=1)
payer_combobox = ttk.Combobox(Frame_info, values=[])
payer_new = Button(Frame_info, text="Ввод нового плательщика или редактирование старого", command=lev_pay)
payer_new.grid(column=2, row=1)
payer_combobox.grid(column=1, row=1)
button_table = Button(Frame_info, command=maketable, text="Составить таблицу")
button_table.grid(row=3, column=2)

in_the_end = Button(root, text="Завершить работу", command=good_ending)
in_the_end.grid(row=4, column=0, columnspan=10)



Button(Frame_info, text="Выбрать словарь CSV", command=table_path_func).grid(column=0, row=0)



root.mainloop()