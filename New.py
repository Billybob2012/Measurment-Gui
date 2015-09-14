import os
from Tkinter import *
import ttk
import time
import datetime
import zipfile
import shutil
import Tix as tk
import webbrowser
from threading import Thread

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
file_ = open("Processes/alternate_name.txt", "w")
file_.close()
file_ = open("Settings/File Locations/OutputFiles.txt", "r")
OutputFolder = file_.readline().rstrip()
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
        ttk.Button(WelcomeFrame, text="View On Git Hub",
                   command=lambda: webbrowser.open("https://github.com/Billybob2012/Measurement-Gui.git")).pack()
        ttk.Button(WelcomeFrame, text="View/Edit Source Code", command=lambda: os.system("notepad.exe New.py")).pack()

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
        settings.insert('', '2','file_locations',text='File Locations')
        settings.insert('file_locations','0','output_folder',text='Output Files Location')
        settings.bind('<<TreeviewSelect>>', self.SettingsDec)

        ### Devices Frame ###
        ttk.Button(DevicesFrame, text='Agilent 34410A', command=lambda: self.Agilent34410AMainMenu()).grid(column=0,
                                                                                                           row=0,
                                                                                                           columnspan=2)

        ### Processes Frame ###
        recipe_list = []
        for file in os.listdir("Recipes/Process_Recipes/"):
            if file.endswith(".txt"):
                recipe_list.append(file[:-4])
        ### Entries ###
        process_name = ttk.Entry(ProcessFrame)
        process_name.grid(column=2, row=6)
        time_delay = ttk.Entry(ProcessFrame)
        time_delay.grid(column=0, row=5)
        ### Buttons ###
        ttk.Button(ProcessFrame, text='Configure Process Que', command=lambda: self.ConfigProcessQue()).grid(column=0,
                                                                                                             row=0)
        ttk.Button(ProcessFrame, text='View/Edit Process Que', command=lambda: self.ViewProcessQue()).grid(column=0,
                                                                                                           row=1)
        ttk.Button(ProcessFrame, text='Clear Process Que', command=lambda: self.ClearProcessQue()).grid(column=0, row=2)
        ttk.Button(ProcessFrame, text='Execute Process Que', command=lambda: self.ExicuteProcessQue()).grid(column=1,
                                                                                                            row=5,
                                                                                                            pady=25)
        ttk.Button(ProcessFrame, text='Save', command=lambda: self.SaveProcessQue(process_name.get())).grid(column=2,
                                                                                                            row=7)
        ttk.Button(ProcessFrame, text='Apply', command=lambda: self.ApplyProcessQueRecipe(process_recipe.get())).grid(
            column=2, row=2)
        ttk.Button(ProcessFrame, text="Add To Que", command=lambda: self.AddToProcessQue(process_recipe.get())).grid(
            column=2, row=3)
        ttk.Button(ProcessFrame, text="Add Delay", command=lambda: self.AddDelayToQue(time_delay.get())).grid(column=0,
                                                                                                              row=6)
        ### Labels ###
        ttk.Label(ProcessFrame, text='Pre-Programed Process Ques').grid(column=2, row=0)
        ttk.Label(ProcessFrame, text='Number of Processes in Que').grid(column=1, row=6)
        ttk.Label(ProcessFrame, text='Save This Que As').grid(column=2, row=5)
        ttk.Label(ProcessFrame, text=str(self.NumberOfProcesses())).grid(row=7, column=1)
        ttk.Label(ProcessFrame, text="Add delay in process que (s)").grid(column=0, row=4)
        ### Combo Boxes ###
        process_recipe = ttk.Combobox(ProcessFrame, values=recipe_list)
        process_recipe.grid(column=2, row=1)

        ### Measurements Frame ###
        ttk.Label(MeasurementFrame, text="Short Term Tests").grid()
        ttk.Button(MeasurementFrame, text='Resistance', command=lambda: self.ResistanceMes()).grid()
        ttk.Button(MeasurementFrame, text='Critical Current', command=lambda: self.CriticalCur()).grid()
        ttk.Button(MeasurementFrame, text='Temperature Vs Resistance', command=lambda: self.TemRes()).grid()
        ttk.Label(MeasurementFrame, text="Long Term Tests").grid(column = 1, row=0)
        ttk.Button(MeasurementFrame, text="Resistance",command=lambda:self.ResistanceMesLongTerm()).grid(column=1, row=1)

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
        ttk.Button(database_button_frame, text='Search Database', command=lambda: self.ViewDatabase()).pack()
        ttk.Button(database_button_frame, text="Backup Database", command=lambda: self.BAckupDatabase()).pack()
        ttk.Button(database_button_frame, text="Clear Database", command=lambda: self.ClearDatabasePrompt()).pack()

    def AddDelayToQue(self, delay):
        global notebook
        _file = open("Processes/process_que.txt", "a")
        _file.write("Wait" + "\n")
        _file.write(delay + "\n")
        _file.close()
        notebook.destroy()
        self.Mainscreen("5")

    def ViewDatabase(self):
        ViewData = Toplevel()

        def SearchDatabase(OP, TM, CN, CT, CI, T, D, TMP):
            def GraphResistance():
                matplotlib.pyplot.ion()
                matplotlib.pyplot.plot(day, resistance)
                matplotlib.pyplot.draw()

            def Statistics():
                def GetAverageResistance():
                    global c
                    c_ = c - 1
                    total = 0
                    divide = c
                    while c_ >= 0:
                        try:
                            total = total + resistance[c_]
                        except:
                            divide -= 1
                        c_ -= 1
                    return total / divide

                def GetLowestRessistance():
                    pass

                def GetHighestResistance():
                    pass

                def GetAverageCriticalCurrent():
                    pass

                StatisticsMenu = Toplevel()
                ttk.Label(StatisticsMenu, text="Average Resistance: " + str(GetAverageResistance())).pack()

            def Delete():
                global c
                c_ = c - 1
                while c_ >= 0:
                    os.remove("Database/" + str(operator[c_]) + "CN" + str(chip_number[c_]) + "TM" + str(
                        measurement_type[c_]) + "CT" + str(chip_type[c_]) + "CI" + str(chip_input[c_]) + "-H" + str(
                        time_[c_])[0:2] + "M" + str(time_[c_])[3:5] + "S" + str(time_[c_])[6:8] + "D" + str(date[c_]))
                    c_ -= 1
                r.destroy()

            def Export():
                def ExportAs(name):
                    global c
                    c_ = c - 1
                    row = 0
                    col = 0
                    row = 0
                    workbook = xlsxwriter.Workbook(str(name).rstrip() + '.xlsx')
                    format = workbook.add_format()
                    format.set_text_wrap()
                    worksheet = workbook.add_worksheet()
                    worksheet.set_column('A:P', 13, format)
                    worksheet.set_row(0, 50, format)
                    worksheet.write(row, col, 'Operator', format)
                    worksheet.write(row, col + 1, 'Time', format)
                    worksheet.write(row, col + 2, 'Date', format)
                    worksheet.write(row, col + 3, 'Chip\nNumber', format)
                    worksheet.write(row, col + 4, 'Chip\nType', format)
                    worksheet.write(row, col + 5, 'Chip\nInput', format)
                    worksheet.write(row, col + 6, 'Chip Temperature'+"\n",format)
                    worksheet.write(row, col + 7, 'Forced', format)
                    worksheet.write(row, col + 8, 'Voltage\nRead', format)
                    worksheet.write(row, col + 9, 'Resistance', format)
                    worksheet.write(row, col + 10, 'Critical\nCurrent', format)
                    worksheet.write(row, col + 11, 'Current\nSteps', format)
                    worksheet.write(row, col + 12, 'Type\nof\nMeasurement', format)
                    row += 1
                    while c_ >= 0:
                        worksheet.write(row, col, operator[c_], format)
                        worksheet.write(row, col + 1, time_[c_], format)
                        worksheet.write(row, col + 2, date[c_], format)
                        worksheet.write(row, col + 3, "=" + str(chip_number[c_]), format)
                        worksheet.write(row, col + 4, chip_type[c_], format)
                        worksheet.write(row, col + 5, "=" + str(chip_input[c_]), format)
                        worksheet.write(row, col + 6, "=" + str(temp[c_])[:], format)
                        worksheet.write(row, col + 7, "=" + str(forced[c_]), format)
                        worksheet.write(row, col + 8, "=" + str(voltage[c_]), format)
                        worksheet.write(row, col + 9, "=" + str(resistance[c_]), format)
                        worksheet.write(row, col + 10, "=" + str(IC[c_]), format)
                        worksheet.write(row, col + 11, "=" + str(current_steps[c_]), format)
                        worksheet.write(row, col + 12, measurement_type[c_], format)
                        row += 1
                        c_ -= 1
                    workbook.close()
                    ExportWindow.destroy()

                ExportWindow = Toplevel()
                ttk.Label(ExportWindow, text="Export Table As").pack()
                excel_name = ttk.Entry(ExportWindow)
                excel_name.pack()
                ttk.Button(ExportWindow, text="Export", command=lambda: ExportAs(excel_name.get())).pack()

            if CI != "CI":
                CI = CI + "-"
            i = 0
            chip_type = []
            input_current_draw = []
            overall_current_draw = []
            operator = []
            chip_number = []
            chip_input = []
            forced = []
            date = []
            time_ = []
            resistance = []
            voltage = []
            temp = []
            IC = []
            current_steps = []
            day = []
            measurement_type = []
            for file in os.listdir("Database/"):
                if file.__contains__(OP) and file.__contains__(TM) and file.__contains__(CN) and file.__contains__(
                        CT) and file.__contains__(
                    CI) and file.__contains__(T) and file.__contains__(D) and file.__contains__(TMP):
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
                    forced.append(forced_)
                    voltage_ = _file_.readline().rstrip()
                    try:
                        voltage.append(float(voltage_))
                    except:
                        voltage.append("")
                    resistance_ = _file_.readline().rstrip()
                    print resistance_
                    try:
                        resistance.append(int(resistance_))
                    except:
                        resistance.append("")
                    temp.append(_file_.readline().rstrip())
                    measurement_type_ = _file_.readline().rstrip()
                    measurement_type.append(measurement_type_)
                    IC_ = _file_.readline().rstrip()
                    IC.append(IC_)
                    current_steps_ = _file_.readline().rstrip()
                    current_steps.append(current_steps_)
                    overall_current_draw.append(_file_.readline().rstrip())
                    input_current_draw.append(_file_.readline().rstrip())
                    i += 1
            r = tk.Tk()
            r.geometry("1275x400")
            r.option_add("*tearOff", False)
            menubar=Menu(r)
            r.config(menu = menubar)
            view = Menu(menubar)
            graph = Menu(menubar)
            delete = Menu(menubar)
            file = Menu(menubar)
            menubar.add_cascade(menu=file, label="File")
            menubar.add_cascade(menu=view, label="View")
            menubar.add_cascade(menu = graph, label = "Graphing")
            file.add_command(label="Delete", command=lambda: Delete())
            file.add_command(label="Export As", command=lambda: Export())
            view.add_command(label="Statistics", command=lambda: Statistics())
            graph.add_command(label = "Graph Resistance Over Time",command=lambda:GraphResistance())
            Results = tk.ScrolledWindow(r, scrollbar=tk.Y)
            Results.pack(fill=tk.BOTH, expand=1)
            tk.Label(Results.window, text="Operator", relief='ridge').grid(column=0, row=0, sticky="WENS")
            tk.Label(Results.window, text="Time", relief='ridge').grid(column=1, row=0, sticky="WENS")
            tk.Label(Results.window, text="Date", relief='ridge').grid(column=2, row=0, sticky="WENS")
            tk.Label(Results.window, text="Chip Number", relief='ridge').grid(column=3, row=0, sticky="WENS")
            tk.Label(Results.window, text="Chip Type", relief='ridge').grid(column=4, row=0, sticky="WENS")
            tk.Label(Results.window, text="Chip Input", relief='ridge').grid(column=5, row=0, sticky="WENS")
            tk.Label(Results.window, text="Chip Temperature", relief='ridge').grid(column=6, row=0, sticky="WENS")
            tk.Label(Results.window, text="Forced", relief='ridge').grid(column=7, row=0, sticky="WENS")
            tk.Label(Results.window, text="Voltage Read", relief='ridge').grid(column=8, row=0, sticky="WENS")
            tk.Label(Results.window, text="Resistance", relief='ridge').grid(column=9, row=0, sticky="WENS")
            tk.Label(Results.window, text="Critical Current (ma)", relief='ridge').grid(column=10, row=0, sticky="WENS")
            tk.Label(Results.window, text="Current Steps (ma)", relief='ridge').grid(column=11, row=0, sticky="WENS")
            tk.Label(Results.window, text="Overall Current Draw", relief='ridge').grid(column=12, row=0, sticky="WENS")
            tk.Label(Results.window, text="Input Current Draw", relief='ridge').grid(column=13, row=0, sticky="WENS")
            tk.Label(Results.window, text="Type Of Measurement", relief='ridge').grid(column=14, row=0, sticky="WENS")
            global c
            c = i
            while i > 0:
                i -= 1
                tk.Label(Results.window, text=operator[i], relief='ridge').grid(column=0, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=time_[i], relief='ridge').grid(column=1, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=date[i], relief='ridge').grid(column=2, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=chip_number[i], relief='ridge').grid(column=3, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=chip_type[i], relief='ridge').grid(column=4, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=chip_input[i], relief='ridge').grid(column=5, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=temp[i], relief='ridge').grid(column=6, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=forced[i], relief='ridge').grid(column=7, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=voltage[i], relief='ridge').grid(column=8, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=resistance[i], relief='ridge').grid(column=9, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=IC[i], relief='ridge').grid(column=10, row=i + 1, sticky="WENS")
                tk.Label(Results.window, text=current_steps[i], relief='ridge').grid(column=11, row=i + 1,
                                                                                     sticky="WENS")
                tk.Label(Results.window, text=overall_current_draw[i], relief='ridge').grid(column=12, row=i + 1,
                                                                                            sticky="WENS")
                tk.Label(Results.window, text=input_current_draw[i], relief='ridge').grid(column=13, row=i + 1,
                                                                                          sticky="WENS")
                tk.Label(Results.window, text=measurement_type[i], relief='ridge').grid(column=14, row=i + 1,
                                                                                        sticky="WENS")
            self.Prggressbar("stop")

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
        ttk.Label(ViewData, text="Chip Tempurature").pack()
        temp = ttk.Combobox(ViewData, values=("Superconducting", "Normal"))
        temp.pack()
        ttk.Label(ViewData, text="Date From").pack()
        date = ttk.Entry(ViewData)
        date.pack()
        ttk.Label(ViewData, text="Date To").pack()
        date_to = ttk.Entry(ViewData)
        date_to.pack()
        ttk.Label(ViewData, text="Time (Hour)").pack()
        time_ = ttk.Entry(ViewData)
        time_.pack()
        ttk.Label(ViewData, text="Input ").pack()
        chip_input = ttk.Entry(ViewData)
        chip_input.pack()
        ttk.Button(ViewData, text="Search",
                   command=lambda: SearchDatabase(operator.get(), "TM" + measurement_type.get(),
                                                  "CN" + chip_number.get(), "CT" + chip_type.get(),
                                                  "CI" + chip_input.get(), "H" + time_.get(), "D" + date.get(),
                                                  "TMP" + temp.get())).pack()

    def ClearDatabasePrompt(self):
        def ClearDatabase():
            for file in os.listdir("Database/"):
                os.remove("Database/" + file)
            CDP.destroy()

        CDP = Toplevel()
        ttk.Label(CDP, text="Are You Sure You Want To Clear The Database?").pack()
        ttk.Button(CDP, text="Yes Clear Database", command=lambda: ClearDatabase()).pack()

    def BAckupDatabase(self):
        def Backup(name):
            zip_ = zipfile.ZipFile(OutputFolder + name + '.zip', 'w')
            zip_.close()
            zip_ = zipfile.ZipFile(OutputFolder + name + '.zip', 'a')
            for file in os.listdir("Database/"):
                print str(file)
                zip_.write("Database/" + str(file))
            zip_.close()
            BD.destroy()

        BD = Toplevel()
        ttk.Label(BD, text="Backup Data Base As").pack()
        backup_name = ttk.Entry(BD)
        backup_name.pack()
        ttk.Button(BD, text="Backup", command=lambda: Backup(backup_name.get())).pack()

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

    def ResistanceMesLongTerm(self):
        global notebook
        def ApplyRecipe():
            pass
        def AddMeasToQue():
            _file_ = open("Processes/process_que.txt", "a")
            _file_.write("Measurement Type: Long Term Resistance" + "\n")
            _file_.write("Operator: " + str(operator.get()) + "\n")
            _file_.write("Type of Chip: " + str(chip_type.get()) + "\n")
            _file_.write(("Chip Number: " + str(chip_number.get()) + "\n"))
            _file_.write(("Super Conducting Voltage: " + str(s_voltage.get()) + "\n"))
            _file_.write(("Normal Conductance Voltage: " + str(r_voltage.get()) + "\n"))
            _file_.write(("Inline Resistor Value: "+str(resistor_value.get())+"\n"))
            _file_.write(("Card Slot Number: "+str(slot_number.get())+"\n"))
            _file_.write(("Input Number: "+str(input_number.get())+"\n"))
            _file_.write("### End Of Measurement ###" + '\n')
            _file_.close()
            ResMenu.destroy()
            notebook.destroy()
            self.Mainscreen('4')
        ResMenu = Toplevel()
        recipe_list = []
        for file in os.listdir("Recipes/ResistanceLongTerm/"):
            if file.endswith(".txt"):
                recipe_list.append(file[:-4])
        ttk.Label(ResMenu, text="Operator").grid()
        operator = ttk.Entry(ResMenu)
        operator.grid()
        ttk.Label(ResMenu, text="Type of Chip").grid()
        chip_type = ttk.Combobox(ResMenu, values=("Lines", "Vias", "Resistors", "JJs"))
        chip_type.grid()
        ttk.Label(ResMenu, text="Chip Number").grid()
        chip_number = ttk.Entry(ResMenu)
        chip_number.grid()
        ttk.Label(ResMenu, text="Super Conducting Voltage").grid()
        s_voltage = ttk.Entry(ResMenu)
        s_voltage.grid()
        ttk.Label(ResMenu, text="Normal Conductance Voltage").grid()
        r_voltage = ttk.Entry(ResMenu)
        r_voltage.grid()
        ttk.Label(ResMenu, text="Inline Resistor Value").grid()
        resistor_value = ttk.Entry(ResMenu)
        resistor_value.grid()
        ttk.Label(ResMenu, text="Slot Number").grid()
        slot_number = ttk.Entry(ResMenu)
        slot_number.grid()
        ttk.Label(ResMenu, text="Input Number").grid()
        input_number = ttk.Entry(ResMenu)
        input_number.grid()
        ttk.Label(ResMenu, text='Choose From a Pre-Prgrmaed Resistance Measurement').grid()
        recipe = ttk.Combobox(ResMenu, values=(recipe_list))
        recipe.grid()
        ttk.Button(ResMenu, text="Apply Recipe", command=lambda: ApplyRecipe(recipe.get())).grid()
        ttk.Label(ResMenu, text='Save This Resistance Measurement As').grid()
        save_as = ttk.Entry(ResMenu)
        save_as.grid()
        ttk.Button(ResMenu, text='Save', command=lambda: SaveRecipe(save_as.get())).grid()
        ttk.Button(ResMenu, text='Add This Measurement To The Que', command=lambda: AddMeasToQue()).grid()

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
        def AddMeasToQue():
            global notebook
            _file_ = open("Processes/process_que.txt", "a")
            _file_.write("Measurement Type: CriticalCurrent" + "\n")
            _file_.write("Operator: " + str(operator.get()) + "\n")
            _file_.write("Type of Chip: " + str(chip_type.get()) + "\n")
            _file_.write(("Chip Number: " + str(chip_number.get()) + "\n"))
            _file_.write("Starting Current: " + str(starting_Current.get()) + '\n')
            _file_.write("Current Steps: " + str(current_steps.get()) + '\n')
            _file_.write("Current Limit: " + str(current_limit.get()) + '\n')
            _file_.write("Voltage Limit: " + str(voltage_limit.get()) + '\n')
            _file_.write("Slot Number: " + str(slot_number.get()) + '\n')
            _file_.write("Input From: " + str(input_from.get()) + '\n')
            _file_.write("Input To: " + str(input_to.get()) + '\n')
            _file_.write("Name of Excel File: " + str(excel_name.get()) + '\n')
            _file_.write("### End Of Measurement ###" + '\n')
            _file_.close()
            CritCur.destroy()
            notebook.destroy()
            self.Mainscreen('4')

        def ApplyRecipe(recipe_):
            _file_ = open("Recipes/CriticalCurrent/" + recipe_ + ".txt", "r")
            chip_type.insert(0, _file_.readline().rstrip())
            starting_Current.insert(0, _file_.readline().rstrip())
            current_steps.insert(0, _file_.readline().rstrip())
            current_limit.insert(0, _file_.readline().rstrip())
            voltage_limit.insert(0, _file_.readline().rstrip())
            slot_number.insert(0, _file_.readline().rstrip())
            input_from.insert(0, _file_.readline().rstrip())
            input_to.insert(0, _file_.readline().rstrip())

        def SaveRecipe(name):
            _file_ = open("Recipes/CriticalCurrent/" + name + ".txt", "w")
            _file_.write(chip_type.get() + "\n")
            _file_.write(starting_Current.get() + "\n")
            _file_.write(current_steps.get() + "\n")
            _file_.write(current_limit.get() + "\n")
            _file_.write(voltage_limit.get() + "\n")
            _file_.write(slot_number.get() + "\n")
            _file_.write(input_to.get() + "\n")
            _file_.write(input_from.get() + "\n")

        recipe_list = []
        for file in os.listdir("Recipes/CriticalCurrent/"):
            if file.endswith(".txt"):
                recipe_list.append(file[:-4])
        CritCur = Toplevel()
        ttk.Label(CritCur, text="Operator").pack()
        operator = ttk.Entry(CritCur)
        operator.pack()
        ttk.Label(CritCur, text="Type of Chip").pack()
        chip_type = ttk.Combobox(CritCur, values=("Lines", "Vias", "Resistors", "JJs"))
        chip_type.pack()
        ttk.Label(CritCur, text="Chip Number").pack()
        chip_number = ttk.Entry(CritCur)
        chip_number.pack()
        ttk.Label(CritCur, text="Starting Current").pack()
        starting_Current = ttk.Entry(CritCur)
        starting_Current.pack()
        ttk.Label(CritCur, text="Current Steps").pack()
        current_steps = ttk.Entry(CritCur)
        current_steps.pack()
        ttk.Label(CritCur, text="Current Limit").pack()
        current_limit = ttk.Entry(CritCur)
        current_limit.pack()
        ttk.Label(CritCur, text="Voltage Limit").pack()
        voltage_limit = ttk.Entry(CritCur)
        voltage_limit.pack()
        ttk.Label(CritCur, text="Slot Number").pack()
        slot_number = ttk.Entry(CritCur)
        slot_number.pack()
        ttk.Label(CritCur, text="Input From").pack()
        input_from = ttk.Entry(CritCur)
        input_from.pack()
        ttk.Label(CritCur, text="Input To").pack()
        input_to = ttk.Entry(CritCur)
        input_to.pack()
        ttk.Label(CritCur, text="Name of Excel File").pack()
        excel_name = ttk.Entry(CritCur)
        excel_name.pack()
        ttk.Label(CritCur, text='Choose From a Pre-Prgrmaed Resistance Measurement').pack()
        recipe = ttk.Combobox(CritCur, values=(recipe_list))
        recipe.pack()
        ttk.Button(CritCur, text="Apply Recipe", command=lambda: ApplyRecipe(recipe.get())).pack()
        ttk.Label(CritCur, text='Save This Resistance Measurement As').pack()
        save_as = ttk.Entry(CritCur)
        save_as.pack()
        ttk.Button(CritCur, text='Save', command=lambda: SaveRecipe(save_as.get())).pack()
        ttk.Button(CritCur, text="Add Process to Que", command=lambda: AddMeasToQue()).pack()

    def TemRes(self):
        TemR = Toplevel()

    def ConfigProcessQue(self):
        ConfigProcess = Toplevel()

        def ChangeZipName(name):
            _file_ = open("Processes/alternate_name.txt", "w")
            _file_.write(name)
            _file_.close()
            ConfigProcess.destroy()
        ttk.Checkbutton(ConfigProcess, text="Include Date and Time of Completion").pack()
        ttk.Label(ConfigProcess, text="Save Zip File As").pack()
        zip_name = ttk.Entry(ConfigProcess)
        zip_name.pack()
        ttk.Button(ConfigProcess, text="Save", command=lambda: ChangeZipName(zip_name.get())).pack()

    def ViewProcessQue(self):
        global notebook
        os.system("notepad.exe Processes/process_que.txt")
        notebook.destroy()
        self.Mainscreen("5")

    def ClearProcessQue(self):
        global notebook
        file_ = open("Processes/process_que.txt", "w")
        file_.close()
        notebook.destroy()
        self.Mainscreen('5')

    def SaveProcessQue(self, name):
        shutil.copy("Processes/process_que.txt", "Recipes/Process_Recipes/")
        os.rename("Recipes/Process_Recipes/process_que.txt", "Recipes/Process_Recipes/" + name + ".txt")

    def ApplyProcessQueRecipe(self, name):
        global notebook
        os.remove("Processes/process_que.txt")
        shutil.copy("Recipes/Process_Recipes/" + name + ".txt", "Processes/")
        os.rename("Processes/" + name + ".txt", "Processes/process_que.txt")
        notebook.destroy()
        self.Mainscreen('5')

    def AddToProcessQue(self, name):
        global notebook
        recipe = open("Recipes/Process_Recipes/" + name + ".txt", "r")
        process_que = open("Processes/process_que.txt", "a")
        copy_line = "null"
        while copy_line != "":
            copy_line = recipe.readline().rstrip()
            check_line = recipe.readline().rstrip()
            if check_line != "":
                process_que.write(copy_line + "\n")
                process_que.write(check_line + "\n")
            else:
                process_que.write((copy_line))
                copy_line = ""
        process_que.close()
        recipe.close()
        notebook.destroy()
        self.Mainscreen("5")
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
        if str(settings.selection()) == "('output_folder',)":
            change_setting = Toplevel()
            ttk.Label(change_setting, text="Folder Location").pack()
            folder_location = ttk.Entry(change_setting)
            folder_location.pack()
            ttk.Button(change_setting, text="Save", command=lambda:save_new_adress("Settings/File Locations/OutputFiles.txt",folder_location.get())).pack()

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

    def Prggressbar(self, command):
        PBAR = Toplevel()
        progress = ttk.Progressbar(PBAR, orient=HORIZONTAL, length=100)
        progress.pack()
        progress.config(mode='indeterminate')
        if command == "start":
            progress.start()
        if command == "stop":
            progress.stop()
            PBAR.destroy()

    def ExicuteProcessQue(self):
        _file_ = open("Processes/process_que.txt", "r")
        type_of_measurement = "T"
        while type_of_measurement != "":
            type_of_measurement = _file_.readline().rstrip()
            if type_of_measurement == "Wait":
                wait_time = _file_.readline().rstrip()
                if wait_time != "u" and wait_time != "U":
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
                    temp = []
                    IC = []
                    current_steps = []
                    day = []
                    measurement_type = []
                    for file in os.listdir("Database/"):
                        if file.__contains__("CN"):
                            _file__ = open("Database/" + file, "r")
                            operator.append(_file__.readline().rstrip())
                            chip_number.append(_file__.readline().rstrip())
                            chip_type.append(_file__.readline().rstrip())
                            chip_input.append(_file__.readline().rstrip())
                            time_.append(_file__.readline().rstrip())
                            date_ = _file__.readline().rstrip()
                            day.append(date_[8:])
                            date.append(date_)
                            forced_ = _file__.readline().rstrip()
                            forced.append(forced_)
                            voltage_ = _file__.readline().rstrip()
                            try:
                                voltage.append(float(voltage_))
                            except:
                                voltage.append("")
                            resistance_ = _file__.readline().rstrip()
                            print resistance_
                            try:
                                resistance.append(int(resistance_))
                            except:
                                resistance.append("")
                            temp.append(_file__.readline().rstrip())
                            measurement_type_ = _file__.readline().rstrip()
                            measurement_type.append(measurement_type_)
                            IC_ = _file__.readline().rstrip()
                            IC.append(IC_)
                            current_steps_ = _file__.readline().rstrip()
                            current_steps.append(current_steps_)
                            i += 1
                    c_ = i - 1
                    row = 0
                    col = 0
                    row = 0
                    workbook = xlsxwriter.Workbook('Database.xlsx')
                    format = workbook.add_format()
                    format.set_text_wrap()
                    worksheet = workbook.add_worksheet()
                    worksheet.set_column('A:P', 13, format)
                    worksheet.set_row(0, 50, format)
                    worksheet.write(row, col, 'Operator', format)
                    worksheet.write(row, col + 1, 'Time', format)
                    worksheet.write(row, col + 2, 'Date', format)
                    worksheet.write(row, col + 3, 'Chip\nNumber', format)
                    worksheet.write(row, col + 4, 'Chip\nType', format)
                    worksheet.write(row, col + 5, 'Chip\nInput', format)
                    worksheet.write(row, col + 6, 'Chip Temperature' + "\n", format)
                    worksheet.write(row, col + 7, 'Forced', format)
                    worksheet.write(row, col + 8, 'Voltage\nRead', format)
                    worksheet.write(row, col + 9, 'Resistance', format)
                    worksheet.write(row, col + 10, 'Critical\nCurrent', format)
                    worksheet.write(row, col + 11, 'Current\nSteps', format)
                    worksheet.write(row, col + 12, 'Type\nof\nMeasurement', format)
                    row += 1
                    while c_ >= 0:
                        print c_
                        worksheet.write(row, col, operator[c_], format)
                        worksheet.write(row, col + 1, time_[c_], format)
                        worksheet.write(row, col + 2, date[c_], format)
                        worksheet.write(row, col + 3, "=" + str(chip_number[c_]), format)
                        worksheet.write(row, col + 4, chip_type[c_], format)
                        worksheet.write(row, col + 5, "=" + str(chip_input[c_]), format)
                        worksheet.write(row, col + 6, "=" + str(temp[c_])[1:], format)
                        worksheet.write(row, col + 7, "=" + str(forced[c_]), format)
                        worksheet.write(row, col + 8, "=" + str(voltage[c_]), format)
                        worksheet.write(row, col + 9, "=" + str(resistance[c_]), format)
                        worksheet.write(row, col + 10, "=" + str(IC[c_]), format)
                        worksheet.write(row, col + 11, "=" + str(current_steps[c_]), format)
                        worksheet.write(row, col + 12, measurement_type[c_], format)
                        row += 1
                        c_ -= 1
                    workbook.close()
                    zip_ = zipfile.ZipFile("DatabaseExcelFile" + '.zip', 'w')
                    zip_.close()
                    zip_ = zipfile.ZipFile("DatabaseExcelFile" + '.zip', 'a')
                    zip_.write('Database.xlsx')
                    zip_.close()
                    for file in os.listdir("Processes/Send_To/"):
                        if file.endswith(".txt"):
                            _file__ = open("Processes/Send_To/" + file, "r")
                            contact_name = _file__.readline().rstrip()
                            contact_email = _file__.readline().rstrip()
                            message = emails.html(
                                html="<p> Greetings: " + contact_name + ",</p>" + "<p>The Auburn Cryo Measurement database was just updated! It was Updated on " + str(
                                    datetime.datetime.now())[:-16] + " at " + str(datetime.datetime.now())[
                                                                              11:-10] + ".</p> <p> War Eagle! </p>",
                                subject="Latest Database",
                                mail_from=("Auburn Cryo Measurement System", "cryomeasurementsystem@gmail.com"))
                            message.attach(data=open("DatabaseExcelFile.zip", 'rb'),
                                           filename="DatabaseExcelFile" + ".zip")
                            r = message.send(to=(contact_name.rstrip(), contact_email), render={"name": "Auburn Cryo"},
                                             smtp={"host": "smtp.gmail.com", "port": 465, "ssl": True,
                                                   "user": "cryomeasurementsystem", "password": "cryoiscold",
                                                   "timeout": 5})
                    time_waited = 0.0
                    prev_state = "NC"
                    while float(wait_time) >= time_waited:
                        time.sleep(0.5)
                        temp_ = float(self.LakeShore336("ask", "KRDG? A"))
                        print temp_
                        if temp_ >= 5.0:
                            forced_voltage = normal_conductance_voltage
                            self.YokogawaGS200('write', 'SOUR:RANG ' + str(forced_voltage))
                            self.YokogawaGS200('write', 'SOUR:LEV ' + str(forced_voltage))
                            prev_state = "NC"
                            time.sleep(1)
                        if prev_state == "NC" and temp_ < 5.0 and temp_cond != "Superconducting":
                            forced_voltage = super_conducting_voltage
                            volt_ramp = float(normal_conductance_voltage)
                            prev_state = "SC"
                            while volt_ramp < float(forced_voltage):
                                volt_ramp += 0.5
                                self.YokogawaGS200('write', 'SOUR:RANG ' + str(volt_ramp))
                                self.YokogawaGS200('write', 'SOUR:LEV ' + str(volt_ramp))
                                time.sleep(0.25)

                        time_waited += 0.5
                else:
                    def WaitScreen():
                        UserWait = Toplevel()
                        ttk.Label(UserWait, text=("Waiting for " + operator + " to continue...")).pack()
                        ttk.Button(UserWait, text="Make Measurement", command=lambda: Continue()).pack()

                        def Continue():
                            UserWait.destroy()
                            self.ExicuteProcessQue()

                    def TempCheck():
                        while 1 == 1:
                            temp_ = float(self.LakeShore336("ask", "KRDG? A"))
                            print temp_
                            if temp_ >= 5.0:
                                forced_voltage = normal_conductance_voltage
                                self.YokogawaGS200('write', 'SOUR:RANG ' + str(forced_voltage))
                                self.YokogawaGS200('write', 'SOUR:LEV ' + str(forced_voltage))
                                prev_state = "NC"
                                time.sleep(1)
                            if prev_state == "NC" and temp_ < 5.0 and temp_cond != "Superconducting":
                                forced_voltage = super_conducting_voltage
                                volt_ramp = float(normal_conductance_voltage)
                                prev_state = "SC"
                                while volt_ramp < float(forced_voltage):
                                    volt_ramp += 0.5
                                    self.YokogawaGS200('write', 'SOUR:RANG ' + str(volt_ramp))
                                    self.YokogawaGS200('write', 'SOUR:LEV ' + str(volt_ramp))
                                    time.sleep(0.25)

                    Thread(target=WaitScreen()).start()
                    Thread(target=TempCheck()).start()

            if type_of_measurement[18:] == "Long Term Resistance":
                operator = _file_.readline().rstrip()[10:]
                chip_type = _file_.readline().rstrip()[14:]
                chip_number = _file_.readline().rstrip()[13:]
                super_conducting_voltage = _file_.readline().rstrip()[26:]
                normal_conductance_voltage = _file_.readline().rstrip()[28:]
                resistor_value = _file_.readline().rstrip()[23:]
                slot_number = _file_.readline().rstrip()[18:]
                input_number = _file_.readline().rstrip()[14:]
                temp = float(self.LakeShore336("ask", "KRDG? A"))
                if temp < 5.0:
                    i = 0
                    forced_voltage = float(super_conducting_voltage)
                    temp_cond = "Superconducting"
                else:
                    forced_voltage = float(normal_conductance_voltage)
                    temp_cond = "Normal"
                i = 0
                self.YokogawaGS200('write', 'SOUR:RANG ' + str(forced_voltage))
                self.YokogawaGS200('write', 'SOUR:LEV ' + str(forced_voltage))
                self.YokogawaGS200('write', 'SENS:REM ON')
                self.YokogawaGS200('write', 'SENS ON')
                self.YokogawaGS200('write', 'OUTP ON')
                self.Keithley7002('write', 'close (@' + slot_number + '!' + input_number + ')')
                time.sleep(0.5)
                self.Keithley7002('write', 'open (@' + slot_number + '!' + "4" + ')')
                self.Keithley7002('write', 'CONF:SLOT' + str(slot_number).rstrip() + ':POLE 2')
                self.YokogawaGS200('write', 'SOUR:FUNC VOLT')
                time.sleep(0.25)
                measured_voltage = self.Agilent34410A('ask', 'MEAS:VOLT:DC?').rstrip()
                self.Keithley7002('write', 'close (@' + slot_number + '!' + "4" + ')')
                time.sleep(0.5)
                self.Keithley7002('write', 'open (@' + slot_number + '!' + input_number + ')')
                dum_resistance = int(
                    ((forced_voltage * float(resistor_value)) / float(measured_voltage)) - float(resistor_value))
                _file__ = open(
                    "Database/" + operator + "CN" + chip_number + "TM" + "Resistance" + "CT" + chip_type + "CI" + str(
                        input_number) + "-" + "H" + str(datetime.datetime.now())[11:-13] + "M" + str(
                        datetime.datetime.now())[14:-10] + "S" + str(datetime.datetime.now())[17:-7] + "D" + str(
                        datetime.datetime.now())[:-16] + "TMP" + temp_cond, "w")
                _file__.write(str(operator) + "\n")
                _file__.write(str(chip_number) + "\n")
                _file__.write(str(chip_type) + "\n")
                _file__.write(str(input_number) + "\n")
                _file__.write(
                    str(datetime.datetime.now())[11:-13] + ":" + str(datetime.datetime.now())[14:-10] + ":" + str(
                        datetime.datetime.now())[17:-7] + "\n")
                _file__.write(str(datetime.datetime.now())[:-16] + "\n")
                _file__.write(str(float(forced_voltage)) + "\n")
                _file__.write(str(float(measured_voltage)) + "\n")
                _file__.write(str(dum_resistance) + "\n")
                _file__.write(str(temp) + "\n")
                _file__.write("Resistance" + "\n")
                _file__.close()
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
                    forced_current = float(starting_current)
                    while float(measured_voltage) < float(voltage_limit) and float(forced_current) < float(
                            current_limit):
                        print str(current_limit) + " " + str(forced_current)
                        self.Keithley7002('write', 'close (@' + slot_number + '!' + input_from + ')')
                        self.Keithley7002('write', 'CONF:SLOT' + str(slot_number).rstrip() + ':POLE 2')
                        self.YokogawaGS200('write', 'SOUR:FUNC CURR')
                        self.YokogawaGS200('write', 'SOUR:RANG ' + str(forced_current))
                        self.YokogawaGS200('write', 'SOUR:LEV ' + str(forced_current))
                        self.YokogawaGS200('write', 'OUTP ON')
                        measured_voltage = self.Agilent34410A('ask', 'MEAS:VOLT:DC?').rstrip()
                        forced_current += float(current_steps)
                    self.YokogawaGS200('write', 'OUTP OFF')
                    self.Keithley7002('write', 'open all')
                    _file__ = open(
                        "Database/" + operator + "CN" + chip_number + "TM" + "CriticalCurrent" + "CT" + chip_type + "CI" + str(
                            input_from) + "-" + "H" + str(datetime.datetime.now())[11:-13] + "M" + str(
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
                    _file__.write("" + "\n")
                    _file__.write("" + "\n")
                    _file__.write("" + "\n")
                    _file__.write("CriticalCurrent" + "\n")
                    _file__.write(str(float(forced_current) * 1000) + "\n")
                    _file__.write(str(float(current_steps) * 1000) + "\n")
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
                time.sleep(3)
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
                        "Database/" + operator + "CN" + chip_number + "TM" + "Resistance" + "CT" + chip_type + "CI" + str(
                            input_from) + "-" + "H" + str(datetime.datetime.now())[11:-13] + "M" + str(
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
                _file = open("Processes/alternate_name.txt", "r")
                process_name = _file.readline().rstrip()
                time_completed = str(datetime.datetime.now())[11:-10]
                date_completed = str(datetime.datetime.now())[:-16]
                if process_name != "":
                    zip_name = str(process_name)
                else:
                    zip_name = ("Date" + str(date_completed) + "Chip" + str(chip_number) + "Type" + str(
                        chip_type) + "Operator" + str(operator))
                    print time_completed
                zip_ = zipfile.ZipFile(OutputFolder + zip_name + '.zip', 'w')
                zip_.close()
                zip_ = zipfile.ZipFile(OutputFolder + zip_name + '.zip', 'a')
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
                        message.attach(data=open(OutputFolder + zip_name + ".zip", 'rb'), filename=zip_name + ".zip")
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
