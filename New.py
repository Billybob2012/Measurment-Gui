# TEST CHANGE
import os
from Tkinter import *
import ttk
import time
import datetime
import zipfile
import shutil
import Tix as tk

import visa
import xlsxwriter
import matplotlib.pyplot
import emails
import emailsms

file_ = open("Settings/Fonts/LabelFontSize.txt", "r")
label_font_size = int(file_.readline().rstrip())
file_.close()
file_ = open("Settings/Fonts/LabelFontColor.txt", "r")
label_font_color = file_.readline().rstrip()
file_.close()
file_ = open("Settings/Fonts/LabelFontType.txt", "r")
label_font_type = file_.readline().rstrip()
file_.close()
file_ = open("Settings/Fonts/ButtonFontSize.txt", "r")
button_font_size = int(file_.readline().rstrip())
file_.close()
file_ = open("Settings/Fonts/ButtonFontColor.txt", "r")
button_font_color = file_.readline().rstrip()
file_.close()
file_ = open("Settings/Fonts/ButtonFontType.txt", "r")
button_font_type = file_.readline().rstrip()
file_.close()
file_ = open("Processes/process_que.txt", "w")
file_.close()
for file in os.listdir("Processes/Send_To/"):
    if file.__contains__(".txt"):
        os.remove(str("Processes/Send_To/" + file))
print str(datetime.datetime.now())
class MainApplication(Frame):
    def Mainscreen(self, place):
        global notebook
        style = ttk.Style()
        style.configure('TButton', foreground=button_font_color, font=(button_font_type, button_font_size))
        style.configure('TLabel', foreground=label_font_color, font=(label_font_type, label_font_size, 'bold'))

        notebook = ttk.Notebook(root)
        notebook.pack()

        ### Frames ###
        WelcomeFrame = ttk.Frame(notebook)
        SettingsFrame = ttk.Frame(notebook)
        DevicesFrame = ttk.Frame(notebook)
        MeasurementFrame = ttk.Frame(notebook)
        ContactsFrame = ttk.Frame(notebook)
        ProcessFrame = ttk.Frame(notebook)
        DatabaseFrame = ttk.Frame(notebook)

        ### Notebook ###
        notebook.add(WelcomeFrame, text='Welcome Screen')
        notebook.add(SettingsFrame, text='Settings')
        notebook.add(ContactsFrame, text='Contacts')
        notebook.add(DevicesFrame, text='Devices')
        notebook.add(MeasurementFrame, text='Measurements')
        notebook.add(ProcessFrame, text='Process Que')
        notebook.add(DatabaseFrame, text='Database')

        notebook.select(int(place))

        ### Welcome Frame ###
        ttk.Label(WelcomeFrame, text='Auburn Cryo Measurement System', style='TLabel').pack()

        ### Settings Frame ###
        global settings
        settings = ttk.Treeview(SettingsFrame)
        settings.pack(side=LEFT)
        settings.insert('', '0', 'devices', text='Devices')
        settings.insert('devices', '0', 'adresses', text='Device Adresses')
        settings.insert('adresses', '0', 'keithley', text='Keithley 7002')
        settings.insert('adresses', '1', 'yokogawa', text='Yokogawa GS200')
        settings.insert('adresses', '2', 'agilent', text='Agilent 34410A')
        settings.insert('adresses', '3', 'lakeshore', text='LakeShore 336')
        settings.insert('', '1', 'fonts', text='Fonts')
        settings.insert('fonts', '0', 'label_fonts', text='Label Fonts')
        settings.insert('label_fonts', '0', 'label_font_size', text='Label Font Size')
        settings.insert('label_fonts', '1', 'label_font_color', text='Label Font Color')
        settings.insert('label_fonts', '2', 'label_font_type', text='Label Font Type')
        settings.insert('fonts', '1', 'button_fonts', text='Button Fonts')
        settings.insert('button_fonts', '0', 'button_font_size', text='Button Font Size')
        settings.insert('button_fonts', '1', 'button_font_color', text='Button Font Color')
        settings.insert('button_fonts', '2', 'button_font_type', text='Button Font Type')
        settings.bind('<<TreeviewSelect>>', self.SettingsDec)

        ### Devices Frame ###
        ttk.Button(DevicesFrame, text='Agilent 34410A', command=lambda: self.Agilent34410AMainMenu()).grid(column=0,
                                                                                                           row=0,
                                                                                                           columnspan=2)

        ### Processes Frame ###
        ### Buttons ###
        ttk.Button(ProcessFrame, text='Configure Process Que', command=lambda: self.ConfigProcessQue()).grid(column=0,
                                                                                                             row=0)
        ttk.Button(ProcessFrame, text='View Process Que').grid(column=0, row=1)
        ttk.Button(ProcessFrame, text='Clear Process Que', command=lambda: self.ClearProcessQue()).grid(column=0, row=2)
        ttk.Button(ProcessFrame, text='Execute Process Que', command=lambda: self.ExicuteProcessQue()).grid(column=1,
                                                                                                            row=5,
                                                                                                            pady=25)
        ttk.Button(ProcessFrame, text='Save').grid(column=2, row=5)
        ### Labels ###
        ttk.Label(ProcessFrame, text='Pre-Programed Process Ques').grid(column=2, row=0)
        ttk.Label(ProcessFrame, text='Number of Processes in Que').grid(column=1, row=6)
        ttk.Label(ProcessFrame, text='Save This Que As').grid(column=2, row=3)
        ttk.Label(ProcessFrame, text=str(self.NumberOfProcesses())).grid(row=7, column=1)
        ### Entries ###
        ttk.Entry(ProcessFrame).grid(column=2, row=4)
        ### Combo Boxes ###
        ttk.Combobox(ProcessFrame).grid(column=2, row=1)

        ### Measurements Frame ###
        ttk.Button(MeasurementFrame, text='Resistance', command=lambda: self.ResistanceMes()).grid()
        ttk.Button(MeasurementFrame, text='Critical Current', command=lambda: self.CriticalCur()).grid()
        ttk.Button(MeasurementFrame, text='Temperature Vs Resistance', command=lambda: self.TemRes()).grid()

        ### Contacts Frame ###
        global cont
        global cont_frame
        cont_frame = ttk.LabelFrame(ContactsFrame, text="Contacts")
        cont_frame.grid(column=0, row=0)
        button_frame = ttk.Frame(ContactsFrame)
        button_frame.grid(row=0, column=1)
        send_to_frame = ttk.LabelFrame(ContactsFrame, text="Send Process Results To")
        send_to_frame.grid(row=0, column=3)
        send_to = ttk.Treeview(send_to_frame)
        cont = ttk.Treeview(cont_frame)
        cont.pack()
        send_to.pack()
        i = 0
        cont.bind('<<TreeviewSelect>>', self.ContactDec)
        for file in os.listdir("Contacts"):
            if file.endswith(".txt"):
                cont.insert('', i, file[:-4], text=file[:-4])
                i += 1
        for file in os.listdir("Processes/Send_To"):
            if file.endswith(".txt"):
                send_to.insert('', i, file[:-4], text=file[:-4])
                i += 1
        ttk.Button(button_frame, text=">", command=lambda: self.SendToContact()).pack()
        ttk.Button(button_frame, text="<").pack()
        ttk.Button(button_frame, text='View / Edit Contact', command=lambda: self.EditContact()).pack()
        ttk.Button(button_frame, text='Delete Contact', command=lambda: self.DeleteContact()).pack()
        ttk.Button(button_frame, text='Add New Contact', command=lambda: self.AddNewContact()).pack()

        ### Database Frame ###

        database_button_frame = ttk.Frame(DatabaseFrame)
        database_button_frame.pack()
        ttk.Button(database_button_frame, text='View Database', command=lambda: self.ViewDatabase()).pack()
        ttk.Button(database_button_frame, text="Backup Database").pack()
        ttk.Button(database_button_frame, text="Clear Database", command=lambda: self.ClearDatabasePrompt()).pack()

    def ViewDatabase(self):
        ViewData = Toplevel()

        def SearchDatabase(OP, TM, CN, CT, CI, T, D):
            def GraphResistance():
                matplotlib.pyplot.ion()
                matplotlib.pyplot.plot(day, resistance)
                matplotlib.pyplot.draw()
            i = 0
            chip_type = []
            operator = []
            chip_number = []
            chip_input = []
            forced = []
            date = []
            time_ = []
            resistance = []
            voltage = []
            IC = []
            current_steps = []
            day = []
            measurement_type = []
            for file in os.listdir("Database/"):
                if file.__contains__(OP) and file.__contains__(TM) and file.__contains__(CN) and file.__contains__(
                        CT) and file.__contains__(
                        CI) and file.__contains__(T) and file.__contains__(D):
                    _file_ = open("Database/" + file, "r")
                    operator.append(_file_.readline().rstrip())
                    chip_number.append(_file_.readline().rstrip())
                    chip_type.append(_file_.readline().rstrip())
                    chip_input.append(_file_.readline().rstrip())
                    time_.append(_file_.readline().rstrip())
                    date_ = _file_.readline().rstrip()
                    day.append(date_[8:])
                    date.append(date_)
                    forced_ = _file_.readline().rstrip()
                    if forced_ == "":
                        forced.append("-")
                    else:
                        forced.append(forced_)
                    voltage_ = _file_.readline().rstrip()
                    if voltage_ == "":
                        voltage.append("-")
                    else:
                        voltage.append(voltage_)
                    resistance_ = _file_.readline().rstrip()
                    if resistance_ == "":
                        resistance.append("-")
                    else:
                        resistance.append(resistance_)
                    measurement_type_ = _file_.readline().rstrip()
                    if measurement_type_ == "":
                        measurement_type.append("-")
                    else:
                        measurement_type.append(measurement_type_)
                    IC_ = _file_.readline().rstrip()
                    if IC_ == "":
                        IC.append("-")
                    else:
                        IC.append(IC_)
                    current_steps_ = _file_.readline().rstrip()
                    if current_steps_ == "":
                        current_steps.append("-")
                    else:
                        current_steps.append(current_steps_)
                    i += 1
            r = tk.Tk()
            Results = tk.ScrolledWindow(r, scrollbar=tk.Y)
            Results.pack(fill=tk.BOTH, expand=1)
            ButtonFrame = tk.Frame(r)
            ButtonFrame.pack(fill="y")
            ttk.Button(ButtonFrame, text="Graph Resistance Over Time", command=lambda: GraphResistance()).grid(column=0,
                                                                                                               row=0)
            ttk.Button(ButtonFrame, text="Percent Difference Of Resistance").grid(column=1, row=0)
            tk.Label(Results.window, text="Operator", relief='ridge').grid(column=0, row=0, sticky="WENS")
            tk.Label(Results.window, text="Time", relief='ridge').grid(column=1, row=0, sticky="WENS")
            tk.Label(Results.window, text="Date", relief='ridge').grid(column=2, row=0, sticky="WENS")
            tk.Label(Results.window, text="Chip Number", relief='ridge').grid(column=3, row=0, sticky="WENS")
            tk.Label(Results.window, text="Chip Type", relief='ridge').grid(column=4, row=0, sticky="WENS")
            tk.Label(Results.window, text="Chip Input", relief='ridge').grid(column=5, row=0, sticky="WENS")
            tk.Label(Results.window, text="Forced", relief='ridge').grid(column=6, row=0, sticky="WENS")
            tk.Label(Results.window, text="Voltage Read", relief='ridge').grid(column=7, row=0, sticky="WENS")
            tk.Label(Results.window, text="Resistance", relief='ridge').grid(column=8, row=0, sticky="WENS")
            tk.Label(Results.window, text="Critical Current", relief='ridge').grid(column=9, row=0, sticky="WENS")
            tk.Label(Results.window, text="Current Steps", relief='ridge').grid(column=10, row=0, sticky="WENS")
            tk.Label(Results.window, text="Type Of Measurement", relief='ridge').grid(column=11, row=0, sticky="WENS")
            while i > 0:
                i -= 1
                tk.Label(Results.window, text=operator[i], relief='ridge').grid(column=0, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=time_[i], relief='ridge').grid(column=1, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=date[i], relief='ridge').grid(column=2, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=chip_number[i], relief='ridge').grid(column=3, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=chip_type[i], relief='ridge').grid(column=4, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=chip_input[i], relief='ridge').grid(column=5, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=forced[i], relief='ridge').grid(column=6, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=voltage[i], relief='ridge').grid(column=7, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=resistance[i], relief='ridge').grid(column=8, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=IC[i], relief='ridge').grid(column=9, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=current_steps[i], relief='ridge').grid(column=10, row=i + 1,
                                                                                     sticky="WENS")
                tk.Label(Results.window, text=measurement_type[i], relief='ridge').grid(column=11, row=i + 1,
                                                                                        sticky="WENS")

        ttk.Label(ViewData, text="Chip Type", relief='groove').pack()
        chip_type = ttk.Combobox(ViewData, values=("Lines", "Vias", "Resistors", "JJs"))
        chip_type.pack()
        ttk.Label(ViewData, text="Operator").pack()
        operator = ttk.Entry(ViewData)
        operator.pack()
        ttk.Label(ViewData, text="Measurement Type").pack()
        measurement_type = ttk.Combobox(ViewData, values=("Resistance", "CriticalCurrent", "TempuratureVsResistance"))
        measurement_type.pack()
        ttk.Label(ViewData, text="Chip Number").pack()
        chip_number = ttk.Entry(ViewData)
        chip_number.pack()
        ttk.Label(ViewData, text="Date").pack()
        date = ttk.Entry(ViewData)
        date.pack()
        ttk.Label(ViewData, text="Time (Hour)").pack()
        time_ = ttk.Entry(ViewData)
        time_.pack()
        ttk.Label(ViewData, text="Input ").pack()
        chip_input = ttk.Entry(ViewData)
        chip_input.pack()
        ttk.Button(ViewData, text="Search",
                   command=lambda: SearchDatabase(operator.get(), "TM" + measurement_type.get(),
                                                  "CN" + chip_number.get(), "CT" + chip_type.get(),
                                                  "CI" + chip_input.get(), "H" + time_.get(), "D" + date.get())).pack()

    def ClearDatabasePrompt(self):
        def ClearDatabase():
            for file in os.listdir("Database/"):
                os.remove("Database/" + file)
            CDP.destroy()

        CDP = Toplevel()
        ttk.Label(CDP, text="Are You Sure You Want To Clear The Database?").pack()
        ttk.Button(CDP, text="Yes Clear Database", command=lambda: ClearDatabase()).pack()

    def ContactDec(self, callback):
        global cont
        global selection
        selection = str(cont.selection())[2:-3]

    def DeleteContact(self):
        global selection
        global notebook
        os.remove("Contacts/" + selection + ".txt")
        notebook.destroy()
        self.Mainscreen('2')

    def AddNewContact(self):
        def SaveContact():
            global notebook
            _file_ = open("Contacts/" + name.get() + ".txt", "w")
            _file_.write(name.get() + '\n')
            _file_.write(email_adress.get() + '\n')
            _file_.write(phone_number.get() + '\n')
            _file_.write(service_provider.get() + '\n')
            _file_.close()
            AddCon.destroy()
            notebook.destroy()
            self.Mainscreen('2')

        AddCon = Toplevel()
        ttk.Label(AddCon, text='Name').grid()
        name = ttk.Entry(AddCon)
        name.grid()
        ttk.Label(AddCon, text='Email').grid()
        email_adress = ttk.Entry(AddCon)
        email_adress.grid()
        ttk.Label(AddCon, text='Phone Number').grid()
        phone_number = ttk.Entry(AddCon)
        phone_number.grid()
        ttk.Label(AddCon, text="Phone Service Provider").grid()
        service_provider = ttk.Combobox(AddCon, values=('Verizon', 'AT&T'))
        service_provider.grid()
        ttk.Button(AddCon, text="Save", command=lambda: SaveContact()).grid()

    def EditContact(self):
        def SaveContact():
            _file_ = open("Contacts/" + selection + ".txt", "w")
            _file_.write(name.get() + '\n')
            _file_.write(email_adress.get() + '\n')
            _file_.write(phone_number.get() + '\n')
            _file_.write(service_provider.get() + '\n')
            _file_.close()
            EditCont.destroy()

        EditCont = Toplevel()
        global selection
        _file_ = open("Contacts/" + selection + ".txt", "r")
        ttk.Label(EditCont, text='Name').grid()
        name = ttk.Entry(EditCont)
        name.insert(0, _file_.readline().rstrip())
        name.grid()
        ttk.Label(EditCont, text='Email').grid()
        email_adress = ttk.Entry(EditCont)
        email_adress.insert(0, _file_.readline().rstrip())
        email_adress.grid()
        ttk.Label(EditCont, text='Phone Number').grid()
        phone_number = ttk.Entry(EditCont)
        phone_number.insert(0, _file_.readline().rstrip())
        phone_number.grid()
        ttk.Label(EditCont, text="Phone Service Provider").grid()
        service_provider = ttk.Combobox(EditCont, values=('Verizon', 'AT&T'))
        service_provider.insert(0, _file_.readline().rstrip())
        service_provider.grid()
        _file_.close()
        ttk.Button(EditCont, text="Save", command=lambda: SaveContact()).grid()

    def SendToContact(self):
        global selection
        global notebook
        shutil.copy("Contacts/" + selection + ".txt", "Processes/Send_To")
        notebook.destroy()
        self.Mainscreen('2')

    def ResistanceMes(self):

        def AddMeasToQue():
            global notebook
            _file_ = open("Processes/process_que.txt", "a")
            _file_.write("Measurement Type: Resistance" + "\n")
            _file_.write("Operator: " + str(operator.get()) + "\n")
            _file_.write("Type of Chip: " + str(chip_type.get()) + "\n")
            _file_.write(("Chip Number: " + str(chip_number.get()) + "\n"))
            _file_.write("Forcing: " + str(forcing.get()) + '\n')
            _file_.write("Amount: " + str(amount_forced.get()) + '\n')
            _file_.write("Card Slot Number: " + str(slot_number.get()) + '\n')
            _file_.write("Card Input From: " + str(input_from.get()) + '\n')
            _file_.write("Card Input To: " + str(input_to.get()) + '\n')
            _file_.write("Excel Name: " + str(name_excel.get()) + '\n')
            _file_.write("### End Of Measurement ###" + '\n')
            _file_.close()
            ResMes.destroy()
            notebook.destroy()
            self.Mainscreen('4')

        def ApplyRecipe(recipe_):
            _file_ = open("Recipes/Resistance/" + recipe_ + ".txt", "r")
            chip_type.insert(0, _file_.readline().rstrip())
            forcing.insert(0, _file_.readline().rstrip())
            amount_forced.insert(0, _file_.readline().rstrip())
            slot_number.insert(0, _file_.readline().rstrip())
            input_from.insert(0, _file_.readline().rstrip())
            input_to.insert(0, _file_.readline().rstrip())

        def SaveRecipe(name):
            _file_ = open("Recipes/Resistance/" + name + ".txt", "w")
            _file_.write(chip_type.get() + "\n")
            _file_.write(forcing.get() + "\n")
            _file_.write(amount_forced.get() + "\n")
            _file_.write(slot_number.get() + "\n")
            _file_.write(input_from.get() + "\n")
            _file_.write(input_to.get() + "\n")

        recipe_list = []
        for file in os.listdir("Recipes/Resistance/"):
            if file.endswith(".txt"):
                recipe_list.append(file[:-4])
        ResMes = Toplevel()
        ttk.Label(ResMes, text="Operator").grid()
        operator = ttk.Entry(ResMes)
        operator.grid()
        ttk.Label(ResMes, text="Type of Chip").grid()
        chip_type = ttk.Combobox(ResMes, values=("Lines", "Vias", "Resistors", "JJs"))
        chip_type.grid()
        ttk.Label(ResMes, text="Chip Number").grid()
        chip_number = ttk.Entry(ResMes)
        chip_number.grid()
        ttk.Label(ResMes, text='Forcing').grid()
        forcing = ttk.Combobox(ResMes, values=('Voltage (V)', 'Current (ma)'))
        forcing.grid()
        ttk.Label(ResMes, text='Amount Forced').grid()
        amount_forced = ttk.Entry(ResMes)
        amount_forced.grid()
        ttk.Label(ResMes, text='Card Slot Number').grid()
        slot_number = ttk.Entry(ResMes)
        slot_number.grid()
        ttk.Label(ResMes, text='Input Number From').grid()
        input_from = ttk.Entry(ResMes)
        input_from.grid()
        ttk.Label(ResMes, text='Input Number To').grid()
        input_to = ttk.Entry(ResMes)
        input_to.grid()
        ttk.Label(ResMes, text='Name of Excel File').grid()
        name_excel = ttk.Entry(ResMes)
        name_excel.grid()
        ttk.Label(ResMes, text='Choose From a Pre-Prgrmaed Resistance Measurement').grid()
        recipe = ttk.Combobox(ResMes, values=(recipe_list))
        recipe.grid()
        ttk.Button(ResMes, text="Apply Recipe", command=lambda: ApplyRecipe(recipe.get())).grid()
        ttk.Label(ResMes, text='Save This Resistance Measurement As').grid()
        save_as = ttk.Entry(ResMes)
        save_as.grid()
        ttk.Button(ResMes, text='Save', command=lambda: SaveRecipe(save_as.get())).grid()
        ttk.Button(ResMes, text='Add This Measurement To The Que', command=lambda: AddMeasToQue()).grid()

    def CriticalCur(self):
        CritCur = Toplevel()
        ttk.Label(CritCur, text="Test").pack()

    def TemRes(self):
        TemR = Toplevel()

    def ConfigProcessQue(self):
        ConfigProcess = Toplevel()
        ttk.Checkbutton(ConfigProcess, text="Include Date and Time of Completion").pack()
        ttk.Label(ConfigProcess, text="Save Zip File As").pack()
        ttk.Entry(ConfigProcess).pack()
        ttk.Button(ConfigProcess, text="Save").pack()

    def ViewProcessQue(self):
        ViewProcess = Toplevel()
        _file_ = open("process_que.txt", "r")
        process_que = _file_.readline().rstrip()
        while process_que != "### End Of Measurement ###":
            pass

    def ClearProcessQue(self):
        global notebook
        file_ = open("Processes/process_que.txt", "w")
        file_.close()
        notebook.destroy()
        self.Mainscreen('5')

    def SettingsDec(self, callback):
        global settings
        global notebook
        style = ttk.Style()
        style.configure('TButton', foreground=button_font_color, font=(button_font_type, button_font_size))
        style.configure('TLabel', foreground=label_font_color, font=(label_font_type, label_font_size))

        def save_new_adress(file, new_adress):
            adress = open(file, 'w')
            adress.write(new_adress)
            adress.close()
            change_setting.destroy()
            notebook.destroy()
            self.Mainscreen('1')

        keithley = StringVar()
        yokogawa = StringVar()
        agilent = StringVar()
        lakeshore = StringVar()
        label_font_size_ = StringVar()
        label_font_color_ = StringVar()
        label_font_type_ = StringVar()
        button_font_size_ = StringVar()
        button_font_color_ = StringVar()
        button_font_type_ = StringVar()
        if str(settings.selection()) == "('keithley',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('Keithley 7002 Address')
            x = open('Settings/DeviceAdresses/Keithley7002.txt', 'r')
            keithley.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=keithley).pack()
            ttk.Button(change_setting, text='Save',
                       command=lambda: save_new_adress('Settings/DeviceAdresses/Keithley7002.txt',
                                                       keithley.get())).pack()
        if str(settings.selection()) == "('agilent',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('Agilent 34410A Address')
            x = open('Settings/DeviceAdresses/Agilent34410A.txt', 'r')
            agilent.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=agilent).pack()
            ttk.Button(change_setting, text='Save',
                       command=lambda: save_new_adress('Settings/DeviceAdresses/Agilent34410A.txt',
                                                       agilent.get())).pack()
        if str(settings.selection()) == "('yokogawa',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('Yokogawa GS200 Address')
            x = open('Settings/DeviceAdresses/YokogawaGS200.txt', 'r')
            yokogawa.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=yokogawa).pack()
            ttk.Button(change_setting, text='Save',
                       command=lambda: save_new_adress('Settings/DeviceAdresses/YokogawaGS200.txt',
                                                       yokogawa.get())).pack()
        if str(settings.selection()) == "('lakeshore',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('LakeShore 336 Address')
            x = open('Settings/DeviceAdresses/LakeShore336.txt', 'r')
            lakeshore.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=lakeshore).pack()
            ttk.Button(change_setting, text='Save',
                       command=lambda: save_new_adress('Settings/DeviceAdresses/LakeShore336.txt',
                                                       lakeshore.get())).pack()
        if str(settings.selection()) == "('label_font_size',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('Label Font Size')
            x = open('Settings/Fonts/LabelFontSize.txt', 'r')
            label_font_size_.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=label_font_size_).pack()
            ttk.Button(change_setting, text='Save', command=lambda: save_new_adress('Settings/Fonts/LabelFontSize.txt',
                                                                                    label_font_size_.get())).pack()
        if str(settings.selection()) == "('label_font_color',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('Label Font Color')
            x = open('Settings/Fonts/LabelFontColor.txt', 'r')
            label_font_color_.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=label_font_color_).pack()
            ttk.Button(change_setting, text='Save', command=lambda: save_new_adress('Settings/Fonts/LabelFontColor.txt',
                                                                                    label_font_color_.get())).pack()
        if str(settings.selection()) == "('label_font_type',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('Label Font Type')
            x = open('Settings/Fonts/LabelFontType.txt', 'r')
            label_font_type_.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=label_font_type_).pack()
            ttk.Button(change_setting, text='Save', command=lambda: save_new_adress('Settings/Fonts/LabelFontType.txt',
                                                                                    label_font_type_.get())).pack()
        if str(settings.selection()) == "('button_font_type',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('Button Font Type')
            x = open('Settings/Fonts/ButtonFontType.txt', 'r')
            button_font_type_.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=button_font_type_).pack()
            ttk.Button(change_setting, text='Save', command=lambda: save_new_adress('Settings/Fonts/ButtonFontType.txt',
                                                                                    button_font_type_.get())).pack()
        if str(settings.selection()) == "('button_font_size',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('button Font Size')
            x = open('Settings/Fonts/buttonFontSize.txt', 'r')
            button_font_size_.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=button_font_size_).pack()
            ttk.Button(change_setting, text='Save', command=lambda: save_new_adress('Settings/Fonts/ButtonFontSize.txt',
                                                                                    button_font_size_.get())).pack()
        if str(settings.selection()) == "('button_font_color',)":
            change_setting = Toplevel()
            change_setting.geometry('300x50')
            change_setting.title('button Font color')
            x = open('Settings/Fonts/buttonFontcolor.txt', 'r')
            button_font_color_.set(x.readline().rstrip())
            ttk.Entry(change_setting, textvariable=button_font_color_).pack()
            ttk.Button(change_setting, text='Save',
                       command=lambda: save_new_adress('Settings/Fonts/ButtonFontColor.txt',
                                                       button_font_color_.get())).pack()

    def Agilent34410AMainMenu(self):
        AgiltentMen = Toplevel()
        style = ttk.Style()
        style.configure('TButton', foreground=button_font_color, font=(button_font_type, button_font_size))
        style.configure('TLabel', foreground=label_font_color, font=(label_font_type, label_font_size))

    def Agilent34410A(self, option, command):
        settings = open('Settings/DeviceAdresses/Agilent34410A.txt', 'r')
        global var
        adress = settings.readline().rstrip()
        print adress
        inst = visa.ResourceManager()
        inst = inst.open_resource(adress)
        if option == 'test':
            var = inst.query(command)
            self.Agilent34410AMeasurementMenu()
        if option == 'write':
            inst.write(command)
        if option == 'ask':
            return inst.query(command)
        inst.close()

    def Keithley7002(self, option, command):
        settings = open('Settings/DeviceAdresses/Keithley7002.txt', 'r')
        adress = settings.readline().rstrip()
        print adress
        inst = visa.ResourceManager()
        inst = inst.open_resource(adress)
        if option == 'write':
            inst.write(command)
        if option == 'ask':
            return inst.query(command)
        inst.close()

    def YokogawaGS200(self, option, command):
        settings = open('Settings/DeviceAdresses/YokogawaGS200.txt', 'r')
        adress = settings.readline().rstrip()
        inst = visa.ResourceManager()
        inst = inst.open_resource(adress)
        if option == 'write':
            inst.write(command)
        if option == 'ask':
            return inst.query(command)
        inst.close()

    def LakeShore336(self, option, command):
        global ans
        global kelv
        settings = open('Settings/DeviceAdresses/LakeShore336.txt', 'r')
        adress = settings.readline().rstrip()
        inst = visa.ResourceManager()
        inst = inst.open_resource(adress)
        if option == 'write':
            inst.write(command)
        if option == 'ask':
            return inst.query(command)
        inst.close()

    def ExicuteProcessQue(self):
        number_of_processes = self.NumberOfProcesses()
        _file_ = open("Processes/process_que.txt", "r")
        while number_of_processes > 0:
            number_of_processes -= 1
            type_of_measurement = _file_.readline().rstrip()
            if type_of_measurement[18:] == "CriticalCurrent":
                operator = _file_.readline().rstrip()[10:]
                chip_type = _file_.readline().rstrip()[14:]
                chip_number = _file_.readline().rstrip()[13:]
                starting_current = _file_.readline().rstrip()[18:]
                starting_current = str(float(starting_current) / 1000)
                current_steps = _file_.readline().rstrip()[15:]
                current_steps = str(float(current_steps) / 1000)
                current_limit = _file_.readline().rstrip()[15:]
                current_limit = str(float(current_limit) / 1000)
                voltage_limit = _file_.readline().rstrip()[15:]
                slot_number = _file_.readline().rstrip()[13:]
                input_from = _file_.readline().rstrip()[12:]
                input_to = _file_.readline().rstrip()[10:]
                excel_name = _file_.readline().rstrip()[20:]
                while int(input_from) < int(input_to) + 1:
                    measured_voltage = '0'
                    while int(measured_voltage) < int(voltage_limit and float(starting_current) < int(current_limit)):
                        self.Keithley7002('write', 'close (@' + slot_number + '!' + input_from + ')')
                        self.Keithley7002('write', 'CONF:SLOT' + str(slot_number).rstrip() + ':POLE 2')

                        self.YokogawaGS200('write', 'SENS:REM ON')
                        self.YokogawaGS200('write', 'SOUR:FUNC CURR')
                        self.YokogawaGS200('write', 'SOUR:RANG ' + str(starting_current))
                        self.YokogawaGS200('write', 'SOUR:LEV ' + str(starting_current))
                        self.YokogawaGS200('write', 'OUTP ON')
                        measured_voltage = self.Agilent34410A('ask', 'MEAS:VOLT:DC?').rstrip()
                        starting_current += int(current_steps)
                    self.YokogawaGS200('write', 'OUTP OFF')
                    self.Keithley7002('write', 'open all')
                    input_from = str(int(input_from) + 1)
            if type_of_measurement[18:] == "Resistance":
                operator = _file_.readline().rstrip()[10:]
                chip_type = _file_.readline().rstrip()[14:]
                chip_number = _file_.readline().rstrip()[13:]
                forcing = _file_.readline().rstrip()[9:]
                forced_amaount = _file_.readline().rstrip()[8:]
                forced_amaount = str(float(forced_amaount.rstrip()) / 1000)
                slot = _file_.readline().rstrip()[18:]
                input_from = _file_.readline().rstrip()[17:]
                input_to = _file_.readline().rstrip()[15:]
                excel_name = _file_.readline().rstrip()[12:]
                _file_.readline()
                workbook = xlsxwriter.Workbook(str(excel_name).rstrip() + '.xlsx')
                format = workbook.add_format()
                format.set_text_wrap()
                worksheet = workbook.add_worksheet()
                self.Keithley7002('write', 'open all')
                self.Keithley7002('write', 'CONF:SLOT' + str(slot).rstrip() + ':POLE 4')
                time.sleep(1)
                row = 0
                col = 0
                worksheet.write(row, col, 'Current', format)
                worksheet.write(row, col + 1, 'Voltage', format)
                worksheet.write(row, col + 2, 'Resistance', format)
                while int(input_from) < int(input_to) + 1:
                    row += 1
                    self.Keithley7002('write', 'close (@' + str(slot).rstrip() + '!' + (str(input_from)).rstrip() + ')')
                    self.YokogawaGS200('write', 'SENS:REM OFF')
                    self.YokogawaGS200('write', 'SOUR:FUNC CURR')
                    self.YokogawaGS200('write', 'SOUR:RANG ' + forced_amaount)
                    self.YokogawaGS200('write', 'SOUR:LEV ' + forced_amaount)
                    self.YokogawaGS200('write', 'OUTP ON')
                    time.sleep(.25)
                    worksheet.write(row, col, '=' + str(float(forced_amaount) * 1000))
                    voltage_meas = str(self.Agilent34410A('ask', 'MEAS:VOLT:DC?'))
                    worksheet.write(row, col + 1, '=' + voltage_meas)
                    worksheet.write(row, col + 2,
                                    '=' + str(int((float(voltage_meas) / float(forced_amaount)))))
                    self.YokogawaGS200('write', 'OUTP OFF')
                    self.Keithley7002('write', 'open all')
                    chart = workbook.add_chart({'type': "column"})
                    chart.add_series({'values': '=Sheet1!$B$2:$B$' + str(row + 1)})
                    worksheet.insert_chart('A7', chart)
                    _file__ = open(
                        "Database/" + operator + "CN" + "TM" + "Resistance" + chip_number + "CT" + chip_type + "CI" + str(
                            input_from) + "H" + str(datetime.datetime.now())[11:-13] + "M" + str(
                            datetime.datetime.now())[14:-10] + "S" + str(datetime.datetime.now())[17:-7] + "D" + str(
                            datetime.datetime.now())[:-16], "w")
                    _file__.write(str(operator) + "\n")
                    _file__.write(str(chip_number) + "\n")
                    _file__.write(str(chip_type) + "\n")
                    _file__.write(str(input_from) + "\n")
                    _file__.write(
                        str(datetime.datetime.now())[11:-13] + ":" + str(datetime.datetime.now())[14:-10] + ":" + str(
                            datetime.datetime.now())[17:-7] + "\n")
                    _file__.write(str(datetime.datetime.now())[:-16] + "\n")
                    _file__.write(str(float(forced_amaount) * 1000) + "\n")
                    _file__.write(str(float(voltage_meas)) + "\n")
                    _file__.write(str(int((float(voltage_meas) / float(forced_amaount)))) + "\n")
                    _file__.write("Resistance" + "\n")
                    input_from = int(input_from) + 1
                workbook.close()
                _file = open("Processes/process_que_settings.txt", "r")
                time_completed = str(datetime.datetime.now())[11:-10]
                date_completed = str(datetime.datetime.now())[:-16]
                zip_name = _file.readline().rstrip() + " " + time_completed + " " + date_completed
                zip_ = zipfile.ZipFile("Output_Files/" + zip_name + '.zip', 'w')
                zip_.close()
                zip_ = zipfile.ZipFile("Output_Files/" + zip_name + '.zip', 'a')
                zip_.write(str(excel_name).rstrip() + '.xlsx')
                os.remove(str(excel_name).rstrip() + '.xlsx')
                zip_.close()
                for file in os.listdir("Processes/Send_To/"):
                    if file.endswith(".txt"):
                        _file_ = open("Processes/Send_To/" + file, "r")
                        contact_name = _file_.readline().rstrip()
                        contact_email = _file_.readline().rstrip()
                        contact_phone_number = _file_.readline().rstrip()
                        contact_service_provider = _file_.readline().rstrip()
                        message = emails.html(
                            html="<p> Greetings: " + contact_name + ",</p>" + "<p>Here are your measurement results, they were completed on " + str(
                                datetime.datetime.now())[:-16] + " at " + str(datetime.datetime.now())[
                                                                          11:-10] + ".</p> <p> War Eagle! </p>",
                            subject=zip_name + " Results",
                            mail_from=("Auburn Cryo Measurement System", "cryomeasurementsystem@gmail.com"))
                        message.attach(data=open("Output_Files/" + zip_name + ".zip", 'rb'), filename=zip_name + ".zip")
                        r = message.send(to=(contact_name.rstrip(), contact_email), render={"name": "Auburn Cryo"},
                                         smtp={"host": "smtp.gmail.com", "port": 465, "ssl": True,
                                               "user": "cryomeasurementsystem", "password": "cryoiscold", "timeout": 5})
                        assert r.status_code == 250
                        smtp = emailsms.gmail_smtp('cryomeasurementsystem@gmail.com', 'cryoiscold')
                        emailsms.send(smtp, contact_phone_number, "Your Measurement Is Completed!",
                                      contact_service_provider)
    def NumberOfProcesses(self):
        _file_ = open("Processes/process_que.txt", "r")
        number_of_processes = 0
        process_read_line = _file_.readline().rstrip()
        while process_read_line != "":
            if process_read_line == "### End Of Measurement ###":
                number_of_processes += 1
            process_read_line = _file_.readline().rstrip()
        _file_.close()
        return number_of_processes

root = Tk()
app = MainApplication(root)
app.Mainscreen('0')
root.mainloop()
