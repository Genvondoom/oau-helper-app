import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

import clipboard
import pandas
from plyer import storagepath

root = Tk()
frame = Frame(root)
frame2 = Frame()
frame3 = Frame(root, bg="grey")
frame4 = Frame(root, bg="white", height=60)
root.minsize(1020, 640)

# root.resizable(True, False)

location = "wefk"
entries = []
dept = ""
options = ['Awaiting Result', 'Token not found', 'Passport Uploaded', 'Wrong result uploaded',
           'No results uploaded', 'Card limit exceeded', 'Invalid card', 'Result not uploaded',
           'Incomplete Result', 'Result not visible', 'Invalid pin', 'Invalid serial',
           'Result checker has been used', 'Pin for Neco not given',
           'Wrong result uploaded', 'Incomplete result',
           'Token linked to another candidate']


# functions
def switch_to_reg():
    global remark_menu
    frame.pack_forget()
    frame3.forget()
    frame4.forget()
    Basic_Mode().start()
    remark_menu.entryconfig("Create In Basic Mode", state="disabled")  # disable menu item
    remark_menu.entryconfig("Create In Advance Mode", state="normal")  # disable menu item
    filemenu.entryconfig("New Department", state="normal")
    filemenu.entryconfig("Load Department", state="normal")

    locmenu.entryconfig("Home", state="normal")  # disable menu item


def switch_to_advance():
    frame.pack_forget()
    frame2.forget()
    remark_menu.entryconfig("Create In Basic Mode", state="normal")  # disable menu item
    remark_menu.entryconfig("Create In Advance Mode", state="disabled")  # disable menu item
    filemenu.entryconfig("New Department", state="normal")
    filemenu.entryconfig("Load Department", state="normal")

    locmenu.entryconfig("Home", state="normal")  # disable menu item
    Advanced_Mode()


def switch_to_home():
    remark_menu.entryconfig("Create In Basic Mode", state="normal")  # disable menu item
    remark_menu.entryconfig("Create In Advance Mode", state="normal")  # disable menu item
    filemenu.entryconfig("New Department", state="disabled")
    filemenu.entryconfig("Load Department", state="disabled")

    locmenu.entryconfig("Home", state="disabled")  # disable menu item
    frame2.forget()
    frame3.forget()
    frame4.forget()
    frame.pack(fill=BOTH, expand=1)
    # t= Advanced_Mode()
    # t.close()


# Menu
menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
locmenu = Menu(menubar, tearoff=0)

remark_menu = Menu(menubar, tearoff=0)
remark_menu.add_command(label="Create In Basic Mode", command=switch_to_reg)
remark_menu.add_command(label="Create In Advance Mode", command=switch_to_advance)
menubar.add_cascade(label="Remark", menu=remark_menu)

menubar.add_cascade(label="Go To", menu=locmenu)
locmenu.add_cascade(label="Home", command=switch_to_home)
locmenu.entryconfig("Home", state="disabled")

menubar.add_cascade(label="File", menu=filemenu)

filemenu.add_command(label="New Department")
filemenu.add_command(label="Load Department")
filemenu.entryconfig("New Department", state="disabled")
filemenu.entryconfig("Load Department", state="disabled")

root.config(menu=menubar)
menubar = Menu(root)

frame.pack(fill=BOTH, expand=1)

label1 = Label(frame, text="Welcome To The Post-Utme\nHelper", font="Calibri 50 bold", justify="center")
label1.place(relx=.5, rely=.3, x=10, y=10, anchor=CENTER)


def download(location):
    test = pandas.read_excel(location, engine='openpyxl')
    rep = test.to_dict('records')
    return rep


new = ""
load = ""


class Basic_Mode:
    def __init__(self):
        self.extra = False
        self.list = []
        self.doc = storagepath.get_documents_dir()
        if os.path.exists(f"{self.doc}\\Helper"):
            pass
        else:
            os.mkdir(f"{self.doc}\\Helper")
        self.doc_loc = f"{self.doc}\\Helper"
        self.frame2 = frame2
        filemenu.entryconfig("New Department", command=self.create_new)
        filemenu.entryconfig("Load Department", command=self.load_existing)

    def save_dept_name(self):
        if self.entry.get():
            global new
            self.dept.configure(text=f"Department: {self.entry.get()}", font="Calibri 40 bold", justify="center")

            new = f"{self.doc_loc}\\{self.entry.get()} basic.xlsx"
            self.entry_win.destroy()
        else:
            messagebox.showerror("Incomplete Info", "Please Enter A Department")
            self.entry.focus()

    def save_entry(self, reg_no, issue, targ):
        entry = {"Reg No": reg_no.upper(), "Remarks": issue.title()}
        if reg_no in targ:
            pass
        else:
            targ.append(entry)

    def create_new(self):
        self.entry_win = Toplevel()
        self.entry_win.geometry("480x60")
        self.entry_win.resizable(False, False)
        self.entry_label = Label(self.entry_win, text="Enter The Department Name:", font="Calibri 15 bold")
        self.entry = Entry(self.entry_win, width=20, font="Calibri 15")
        self.btn_save = Button(self.entry_win, text="save", width=15, command=self.save_dept_name)
        self.entry_label.grid(row=0, column=0)
        self.entry.focus()
        self.entry.grid(row=0, column=1)
        self.btn_save.grid(row=1, column=1)

    def load_existing(self):
        global load
        load = filedialog.askopenfilename(initialdir="C:/", title="Select a file",
                                          filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        test = self.doc.strip("Documents").lstrip("C:\\").lstrip("Users").strip("\\")

        name = load.strip(f"C:/Users{test}/").rstrip("basic.xlsx").strip("Documents/Helper/")

        self.dept.configure(text=f"Department: {name}", font="Calibri 40 bold", justify="center")
        self.list = download(load)
        print("done")

    def additional_remark(self):
        self.extra = True
        self.info_3 = Label(self.frame2, text="Enter second exam's remark", font="Calibri 20")
        self.info_3.place(relx=.1, rely=.6)
        self.remark_2 = ttk.Combobox(self.frame2, value=options, width=35, font="Calibri 20", justify="left")
        self.remark_2.place(relx=.45, rely=.6)
        self.btn1.place_forget()

    def check_duplicates(self, list):
        for x in list:
            if x['Reg No'] == f"{self.reg_no.get()}".upper():
                return True

    def save(self):

        if new or load:
            if self.reg_no.get() and self.remark.get():
                if new:
                    if os.path.exists(new):
                        report = download(new)

                        if self.check_duplicates(report):
                            messagebox.showwarning("Duplicated Entry",
                                                   f"Sorry {self.reg_no.get()} Already Exists")
                        else:
                            if self.remark_2.get():
                                self.save_entry(self.reg_no.get(), f"{self.remark.get()} and {self.remark_2.get()}",
                                                report)
                                self.remark_2.destroy()
                                self.info_3.destroy()
                                btn1 = Button(self.frame2, text="Add Remark", width=30, height=2,
                                              command=self.additional_remark)
                                btn1.place(relx=.6, rely=.6)
                            else:
                                self.save_entry(self.reg_no.get(), self.remark.get(), report)
                            df = pandas.DataFrame.from_dict(report)
                            df.to_excel(new, index=False)
                            self.clear()
                    else:
                        report = []
                        if self.remark_2.get():
                            self.save_entry(self.reg_no.get(), f"{self.remark.get()} and {self.remark_2.get()}", report)
                            self.remark_2.forget()
                            self.info_3.forget()
                            btn1 = Button(self.frame2, text="Add Remark", width=30, height=2,
                                          command=self.additional_remark)
                            btn1.place(relx=.6, rely=.6)
                        else:
                            self.save_entry(self.reg_no.get(), self.remark.get(), report)
                        df = pandas.DataFrame(report)
                        df.to_excel(new, index=False)
                        self.clear()



                elif load:
                    report = download(load)

                    if self.check_duplicates(report):
                        messagebox.showwarning("Duplicated Entry",
                                               f"Sorry {self.reg_no.get()} Already Exists")
                    else:
                        if self.remark_2.get():
                            self.save_entry(self.reg_no.get(), f"{self.remark.get()} and {self.remark_2.get()}",
                                            report)
                            self.remark_2.forget()
                            self.info_3.forget()
                            btn1 = Button(self.frame2, text="Add Remark", width=30, height=2,
                                          command=self.additional_remark)
                            btn1.place(relx=.6, rely=.6)
                        else:
                            self.save_entry(self.reg_no.get(), self.remark.get(), report)

                        df = pandas.DataFrame.from_dict(report)
                        df.to_excel(load, index=False)
                        self.clear()

            else:
                messagebox.showwarning("Incomplete Info",
                                       "Please fill the registration number and at least the first remark!!")
        else:
            messagebox.showwarning("No Save Directory",
                                   "Please Select New Or Load InThe File Menu To Determine Where "
                                   "Your Entries Will Be Saved")

    def clear(self):
        self.remark.delete(0, END)
        if self.extra is True:
            self.remark_2.destroy()
        self.reg_no.delete(0, END)
        try:
            self.info_3.destroy()
            self.remark_2.destroy()

            self.btn1.place(relx=.6, rely=.6)
        except AttributeError:
            pass

    def start(self):

        self.frame2.pack(expand=1, fill=BOTH)
        self.dept = Label(self.frame2, text="Department: Please select one", font="Calibri 40 bold", justify="center")
        self.dept.place(relx=.1, rely=.1)
        self.info_1 = Label(self.frame2, text="Enter registration number", font="Calibri 20", justify="center")
        self.info_1.place(relx=.1, rely=.3)
        self.reg_no = Entry(self.frame2, width=36, font="Calibri 20", justify="center")
        self.reg_no.place(relx=.45, rely=.3)

        self.info_2 = Label(self.frame2, text="Enter first exam's remark", font="Calibri 20", justify="center")
        self.info_2.place(relx=.1, rely=.44)

        self.remark = ttk.Combobox(self.frame2, value=options, width=35, font="Calibri 20", justify="left")
        self.remark.place(relx=.45, rely=.45)

        self.btn1 = Button(self.frame2, text="Add Remark", width=30, height=2, command=self.additional_remark)
        self.btn1.place(relx=.6, rely=.6)

        self.btn_2 = Button(self.frame2, text="Add Entry", width=30, height=2, command=self.save)
        self.btn_2.place(relx=.6, rely=.82)

        # self.remark_2 = ttk.Combobox(self.frame2, value=options, width=35, font="Calibri 20", justify="left")

    def close(self):
        self.frame2.grid_forget()
        filemenu.entryconfig("New Department", state="disabled")  # disable menu item


class Advanced_Mode:
    def __init__(self):
        self.remark_list = []
        self.current = 0
        self.doc = storagepath.get_documents_dir()
        if os.path.exists(f"{self.doc}\\Helper"):
            pass
        else:
            os.mkdir(f"{self.doc}\\Helper")
        self.doc_loc = f"{self.doc}\\Helper"
        self.framer()
        # self.jump_start()
        self.new = False
        self.lowd = False
        filemenu.entryconfig("New Department", command=self.jump_start)
        filemenu.entryconfig("Load Department", command=self.load_start)

    def check(self):
        if len(self.listeh) == self.current + 1:
            self.btn_next.configure(state="disabled")

        elif self.current == 0:
            self.btn_prev.configure(state="disabled")

    def open_file(self):
        file = filedialog.askopenfilename(initialdir="C:/", title="Select Master File",
                                          filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.master_file_entry.delete(0, END)
        self.master_file_entry.insert(0, file)
        self.loader.lift(aboveThis=root)

    def load_open(self):
        self.location_3 = filedialog.askopenfilename(initialdir=self.doc_loc,
                                                     title="Select Your Already Existing File",
                                                     filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.file_name.insert(0, self.location_3)
        self.loader.lift(aboveThis=root)

    def save_start_new(self):

        self.location = self.master_file_entry.get()
        file = self.file_name.get()
        start_point = self.start.get()

        if self.new is True:
            self.location_2 = f"{self.doc_loc}\\{file} advanced.xlsx"
        elif self.lowd is True:
            self.location_2 = file

        if self.location or self.location_2:
            test = pandas.read_excel(self.location, engine="openpyxl")
            self.listeh = test.to_dict('records')

            if start_point:
                for x in range(len(self.listeh)):
                    if self.listeh[x]['REGNO'] == start_point:
                        self.current = int(x)
        self.clear()

        self.load(self.listeh, self.current)
        self.loader.destroy()
        self.check()
        self.new = False
        self.lowd = False

    def next(self):
        if len(self.listeh) != self.current + 1:
            self.current = self.current + 1
            self.clear()
            self.load(self.listeh, self.current)
            self.btn_prev.configure(state="normal")

        else:
            self.btn_next.configure(state="disabled")

    def prev(self):
        if self.current > 0:
            self.current = self.current - 1
            self.clear()
            self.load(self.listeh, self.current)
            self.btn_next.configure(state="normal")
        if self.current == 0:
            self.btn_prev.configure(state="disabled")

    def jump_start(self):
        self.loader = Toplevel()
        self.loader.lift(aboveThis=root)
        self.master_file_label = Label(self.loader, text="Select Your Main Department File", font="Calibri 20")
        self.master_file_label.grid(row=0, column=0)
        self.master_file_entry = Entry(self.loader, width=20, font="Calibri 20")
        self.master_file_entry.grid(row=0, column=1)
        self.select_master = Button(self.loader, text="Select", width=10, command=self.open_file)
        self.select_master.grid(row=0, column=2)

        self.file_name_lbl = Label(self.loader, text="Input The Department Name", font="Calibri 20")
        self.file_name_lbl.grid(row=1, column=0)
        self.file_name = Entry(self.loader, width=20, font="Calibri 20 bold")
        self.file_name.grid(row=1, column=1)

        self.start_point_lbl = Label(self.loader, text="Enter Starting Point", font="Calibri 20")
        self.start_point_lbl.grid(row=2, column=0)
        self.start = Entry(self.loader, width=20, font="Calibri 20")
        self.start.grid(row=2, column=1)

        self.save = Button(self.loader, text="Save", width=10, command=self.save_start_new)
        self.save.grid(row=3, column=1)
        self.new = True

    def load_start(self):
        self.loader = Toplevel()
        self.loader.lift(aboveThis=root)
        self.master_file_label = Label(self.loader, text="Select Your Main Department File", font="Calibri 20")
        self.master_file_label.grid(row=0, column=0)
        self.master_file_entry = Entry(self.loader, width=20, font="Calibri 20")
        self.master_file_entry.grid(row=0, column=1)
        self.select_master = Button(self.loader, text="Select", width=10, command=self.open_file)
        self.select_master.grid(row=0, column=2)

        self.file_name_lbl = Label(self.loader, text="Input The Department Name", font="Calibri 20")
        self.file_name_lbl.grid(row=1, column=0)
        self.file_name = Entry(self.loader, width=20, font="Calibri 20 bold")
        self.file_name.grid(row=1, column=1)
        self.select_slave = Button(self.loader, text="Select", width=10, command=self.load_open)
        self.select_slave.grid(row=1, column=2)

        self.start_point_lbl = Label(self.loader, text="Enter Starting Point", font="Calibri 20")
        self.start_point_lbl.grid(row=2, column=0)
        self.start = Entry(self.loader, width=20, font="Calibri 20")
        self.start.grid(row=2, column=1)

        self.save = Button(self.loader, text="Save", width=10, command=self.save_start_new)
        self.save.grid(row=3, column=1)
        self.lowd = True

    def framer(self):
        self.frame2 = frame3
        self.frame2.pack(fill=BOTH, expand=1)

        self.regno_lbl = Label(self.frame2, text="Registration Number", font="Calibri 20 bold", justify="left",
                               bg="grey")
        self.regno_lbl.place(relx=.05, rely=.15)
        self.reg_no = Entry(self.frame2, width=25, font="Calibri 20 bold", justify="left", )
        self.reg_no.place(relx=.35, rely=.15)

        # Widgets for first label
        self.first_lbl = Label(self.frame2, text="First result", font="Calibri 20 bold", justify="left", bg="blue")
        self.first_lbl.place(relx=.05, rely=.23)

        self.serial_lbl = Label(self.frame2, text="Serial Number", font="Calibri 20 bold", justify="left", bg="grey")
        self.serial_lbl.place(relx=.05, rely=.304)
        self.serial = Entry(self.frame2, width=25, font="Calibri 20 bold", justify="left", )
        self.serial.place(relx=.35, rely=.304)
        self.copy1 = Button(self.frame2, text="Copy", width=10, command=lambda: self.copy(self.serial.get()))
        self.copy1.place(relx=.75, rely=.304)

        self.pin_lbl = Label(self.frame2, text="Pin", font="Calibri 20 bold", justify="left", bg="grey")
        self.pin_lbl.place(relx=.05, rely=.37)
        self.pin = Entry(self.frame2, width=25, font="Calibri 20 bold", justify="left", )
        self.pin.place(relx=.35, rely=.37)
        self.copy1 = Button(self.frame2, text="Copy", width=10, command=lambda: self.copy(self.pin.get()))
        self.copy1.place(relx=.75, rely=.37)

        self.serial_lbl = Label(self.frame2, text="Exam Number", font="Calibri 20 bold", justify="left", bg="grey")
        self.serial_lbl.place(relx=.05, rely=.436)
        self.serial_no = Entry(self.frame2, width=25, font="Calibri 20 bold", justify="left", )
        self.serial_no.place(relx=.35, rely=.436)
        self.copy1 = Button(self.frame2, text="Copy", width=10, command=lambda: self.copy(self.serial_no.get()))
        self.copy1.place(relx=.75, rely=.436)

        self.remark_lbl = Label(self.frame2, text="Remark", font="Calibri 20 bold", justify="left", bg="grey")
        self.remark_lbl.place(relx=.05, rely=.502)
        self.remark1 = ttk.Combobox(self.frame2, value=options, width=30, font="Calibri 15", justify="left")
        self.remark1.place(relx=.35, rely=.502)

        # List for sec label
        self.sec_lbl = Label(self.frame2, text="Second Result", font="Calibri 20 bold", justify="left", bg="red")
        self.sec_lbl.place(relx=.05, rely=.59)

        self.serial_no_lbl = Label(self.frame2, text="Serial Number", font="Calibri 20 bold", justify="left",
                                   bg="grey")
        self.serial_no_lbl.place(relx=.05, rely=.65)
        self.serial_no2 = Entry(self.frame2, width=25, font="Calibri 20 bold", justify="left", )
        self.serial_no2.place(relx=.35, rely=.65)
        self.copy1 = Button(self.frame2, text="Copy", width=10, command=lambda: self.copy(self.serial_no2.get()))
        self.copy1.place(relx=.75, rely=.65)

        self.pin_2_lbl = Label(self.frame2, text="Pin", font="Calibri 20 bold", justify="left", bg="grey")
        self.pin_2_lbl.place(relx=.05, rely=.717)
        self.pin2 = Entry(self.frame2, width=25, font="Calibri 20 bold", justify="left", )
        self.pin2.place(relx=.35, rely=.717)
        self.copy1 = Button(self.frame2, text="Copy", width=10, command=lambda: self.copy(self.pin2.get()))
        self.copy1.place(relx=.75, rely=.717)

        self.exam_no2_lbl = Label(self.frame2, text="Exam Number", font="Calibri 20 bold", justify="left", bg="grey")
        self.exam_no2_lbl.place(relx=.05, rely=.785)
        self.exam_no2 = Entry(self.frame2, width=25, font="Calibri 20 bold", justify="left", )
        self.exam_no2.place(relx=.35, rely=.785)
        self.copy1 = Button(self.frame2, text="Copy", width=10, command=lambda: self.copy(self.exam_no2.get()))
        self.copy1.place(relx=.75, rely=.785)

        self.remark_lbl = Label(self.frame2, text="Remark 2", font="Calibri 20 bold", justify="left", bg="grey")
        self.remark_lbl.place(relx=.05, rely=.85)
        self.remark2 = ttk.Combobox(self.frame2, value=options, width=30, font="Calibri 15", justify="left")
        self.remark2.place(relx=.35, rely=.85)

        self.frame3 = frame4
        self.frame3.pack(fill=X)

        self.btn_prev = Button(self.frame3, text="Previous", width=25, height=2, command=self.prev)
        self.btn_prev.place(relx=.05, rely=.15)

        self.btn_save = Button(self.frame3, text="Save", width=25, height=2, command=self.save_remarks)
        self.btn_save.place(relx=.4, rely=.15)

        self.btn_next = Button(self.frame3, text="Next", width=25, height=2, command=self.next)
        self.btn_next.place(relx=.77, rely=.15)

    def clear(self):
        self.reg_no.delete(0, END)
        self.pin.delete(0, END)
        self.serial.delete(0, END)
        self.serial_no.delete(0, END)
        self.serial_no2.delete(0, END)
        self.pin2.delete(0, END)
        self.exam_no2.delete(0, END)
        self.remark2.delete(0, END)
        self.remark1.delete(0, END)

    def copy(self, text):
        clipboard.copy(text)

    def load(self, liste, targ):
        try:
            self.dept_label = Label(self.frame2, text=f"Department: {liste[targ]['COURSE']}".title(),
                                    font="Calibri 30 bold", justify="left",
                                    bg="grey")
            self.dept_label.place(relx=.05, rely=.05, )

            self.first_year = Label(self.frame2, text=f"Year: {liste[targ]['EXAM TYPE(1) YEAR']}",
                                    font="Calibri 20 bold",
                                    justify="left",
                                    bg="grey")
            self.first_year.place(relx=.75, rely=.23)

            self.reg_no.insert(0, str(liste[targ]['REGNO']))
            # self.reg_no.configure(state="readonly")

            self.pin.insert(0, liste[targ]['EXAM TYPE(1) PIN'])
            self.serial.insert(0, liste[targ]["EXAM TYPE(1)  Serial No"])
            self.serial_no.insert(0, liste[targ]["EXAM TYPE(1) NO"])

            sec_year = Label(self.frame2, text=f"Year: {liste[targ]['EXAM TYPE(2) YEAR']}", font="Calibri 20 bold",
                             justify="left",
                             bg="grey")
            sec_year.place(relx=.75, rely=.58)

            self.serial_no2.insert(0, liste[targ]["EXAM TYPE(2) Serial No"])
            self.pin2.insert(0, liste[targ]["EXAM TYPE(2) PIN"])
            self.exam_no2.insert(0, liste[targ]["EXAM TYPE(2) NO"])

        except KeyError:
            print("please remove all spaces before or after your headers in the excel file")

    def check_duplicates(self, list):
        for x in list:
            if x['REGNO'] == f"{self.reg_no.get()}".upper():
                return True

    def save_remarks(self):
        if self.remark1.get() and self.remark2.get():
            self.listeh[self.current]['REMARK '] = f"{self.remark1.get()}, {self.remark2.get()}"

        elif self.remark1.get():
            self.listeh[self.current]['REMARK'] = self.remark1.get()
        elif self.remark2.get():
            self.listeh[self.current]["REMARK"] = self.remark2.get()
        if os.path.exists(self.location_2):

            self.remark_list = []
            test = pandas.read_excel(self.location_2, engine="openpyxl")
            self.remark_list = test.to_dict('records')

        self.remark_list.append(self.listeh[self.current])
        self.upload(self.remark_list)

    def upload(self, lists):
        if self.location_2:
            if os.path.exists(self.location_2):
                df = pandas.DataFrame.from_dict(lists)
                df.to_excel(self.location_2, index=False)

            else:
                df = pandas.DataFrame(lists)

                df.to_excel(self.location_2, index=False)

        elif self.location_3:
            df = pandas.DataFrame.from_dict(lists)
            df.to_excel(self.location_2, index=False)

    def close(self):
        self.frame2.destroy()
        self.frame3.destroy()


mainloop()
