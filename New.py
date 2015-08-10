# TEST CHANGE
import os
from Tkinter import *
import ttk

import visa

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
file_ = open("process_que.txt", "w")
file_.close()


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

        ### Notebook ###
        notebook.add(WelcomeFrame, text='Welcome Screen')
        notebook.add(SettingsFrame, text='Settings')
        notebook.add(ContactsFrame, text='Contacts')
        notebook.add(DevicesFrame, text='Devices')
        notebook.add(MeasurementFrame, text='Measurements')
        notebook.add(ProcessFrame, text='Process Que')

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
        ttk.Button(ProcessFrame, text='Clear Process Que').grid(column=0, row=2)
        ttk.Button(ProcessFrame, text='Execute Process Que').grid(column=1, row=5, pady=25)
        ttk.Button(ProcessFrame, text='Save').grid(column=2, row=5)
        ### Labels ###
        ttk.Label(ProcessFrame, text='Pre-Programed Process Ques').grid(column=2, row=0)
        ttk.Label(ProcessFrame, text='Number of Processes in Que').grid(column=1, row=6)
        ttk.Label(ProcessFrame, text='Save This Que As').grid(column=2, row=3)
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
        cont_frame = ttk.Frame(ContactsFrame)
        cont_frame.grid(column=0, row=0)
        button_frame = ttk.Frame(ContactsFrame)
        button_frame.grid(row=0, column=1)
        cont = ttk.Treeview(cont_frame)
        cont.pack()
        i = 0
        cont.bind('<<TreeviewSelect>>', self.ContactDec)
        for file in os.listdir("Contacts"):
            if file.endswith(".txt"):
                cont.insert('', i, file[:-4], text=file[:-4])
                i += 1
        ttk.Button(button_frame, text='View / Edit Contact', command=lambda: self.EditContact()).pack()
        ttk.Button(button_frame, text='Delete Contact', command=lambda: self.DeleteContact()).pack()
        ttk.Button(button_frame, text='Add New Contact', command=lambda: self.AddNewContact()).pack()

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

    def ResistanceMes(self):
        def AddMeasToQue():
            _file_ = open("process_que.txt", "a")
            _file_.write("### Start Of Measurement ###" + '\n')
            _file_.write("Measurement Type: Resistance" + "\n")
            _file_.write("Forcing: " + str(forcing.get()) + '\n')
            _file_.write("Amount: " + str(amount_forced.get()) + '\n')
            _file_.write("Card Slot Number: " + str(slot_number.get()) + '\n')
            _file_.write("Card Input From: " + str(input_from.get()) + '\n')
            _file_.write("Card Input To: " + str(input_to.get()) + '\n')
            _file_.write("Excel Name: " + str(name_excel.get()) + '\n')
            _file_.write("### End Of Measurement ###" + '\n')
            _file_.close()
            ResMes.destroy()

        ResMes = Toplevel()
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
        recipe = ttk.Combobox(ResMes)
        recipe.grid()
        ttk.Label(ResMes, text='Save This Resistance Measurement As').grid()
        save_as = ttk.Entry(ResMes)
        save_as.grid()
        ttk.Button(ResMes, text='Save').grid()
        ttk.Button(ResMes, text='Add This Measurement To The Que', command=lambda: AddMeasToQue()).grid()

    def CriticalCur(self):
        CritCur = Toplevel()
        ttk.Label(CritCur, text="Test").pack()

    def TemRes(self):
        TemR = Toplevel()

    def ConfigProcessQue(self):
        ConfigProcess = Toplevel()

    def ViewProcessQue(self):
        ViewProcess = Toplevel()
        _file_ = open("process_que.txt", "r")
        process_que = _file_.readline().rstrip()
        while process_que != "### End Of Measurement ###":
            pass

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


root = Tk()
# root.geometry('640x480')
app = MainApplication(root)
app.Mainscreen('0')
root.mainloop()
