import winsound

try:
    import visa
except ImportError:
    print "Please install PyVisa Library"
from Tkinter import *
import time

try:
    import xlsxwriter
except:
    print 'Please install XlsxWriter'
try:
    import serial
except:
    print 'Please install PySerial'
try:
    import matplotlib.pyplot
except:
    print 'Please install MatPlotLib'
try:
    open('4_Wire_Recipes.txt', 'r')
except:
    x = open('4_Wire_Recipes.txt', 'w')
    x.write('None' + '\n')
    x.close()
    x = open('None.txt', 'w')
    x.close()
    print 'Made 4_Wire_Recipes.txt File'
try:
    open('V_Vs_C_Recipes.txt', 'r')
except:
    x = open('V_Vs_C_Recipes.txt', 'w')
    x.write('None' + '\n')
    x.close()
    x = open('None.txt', 'w')
    x.close()
try:
    open('Process_Recipes.txt', 'r')
except:
    x = open('Process_Recipes.txt', 'w')
    x.write('None' + '\n')
    x.close()
    x = open('None.txt', 'w')
    x.close()

kelv = 0
ans = '0'


class Application(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        self.grid()
        self.DeviceMen()

    def DeviceMen(self):
        root.wm_state('zoomed')
        global count
        global var
        global forced
        global range
        global input
        global output
        global name
        global measure
        global to
        global fr
        global tm
        global rate
        global wanted_temp
        global graph
        global slot
        global outp
        global inp
        global recipe
        global recipe_name
        global back_img
        global add_process_img
        global save_img
        global process_recipe
        global process_recipe_names
        global agilent_img
        global keithley_img
        global yokogawa_img
        global lakeshore_img
        global apply_img
        global exicute_img
        global font_size
        font_size = 15
        exicute_img = PhotoImage(file="exicute.gif")
        apply_img = PhotoImage(file="apply.gif")
        lakeshore_img = PhotoImage(file="lakeshore.gif")
        yokogawa_img = PhotoImage(file="yokogawa.gif")
        keithley_img = PhotoImage(file="keithley.gif")
        agilent_img = PhotoImage(file="agilent.gif")
        process_recipe = StringVar()
        process_recipe_names = StringVar()
        save_img = PhotoImage(file="save.gif")
        back_img = PhotoImage(file="back.gif")
        add_process_img = PhotoImage(file="add_to_que.gif")
        recipe_name = StringVar()
        recipe = StringVar()
        inp = StringVar()
        outp = StringVar()
        graph = StringVar()
        range = StringVar()
        measure = StringVar()
        forced = StringVar()
        fr = StringVar()
        to = StringVar()
        name = StringVar()
        tm = StringVar()
        rate = StringVar()
        wanted_temp = StringVar()
        slot = StringVar()
        var = '0'
        count = 0
        process = open('process_que.txt', 'w')
        process.close()  # closes the file
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text='Select a device to connect to', font=(font_size)).grid(columnspan=2)
        Button(self, padx=25, pady=25, text='Agilent 34410A DMM', font=(font_size), image=agilent_img, compound=LEFT,
               command=lambda: self.Agilent34410AMainMenu()).grid(column=1, row=1)
        Button(self, padx=25, pady=25, text='Keithley 7002 Switching Machine', font=(font_size), image=keithley_img,
               compound=LEFT,
               command=lambda: self.Keithley7002MainMenu()).grid(column=0, row=1)
        Button(self, padx=25, pady=25, text='Yokogawa GS200', font=(font_size), image=yokogawa_img, compound=LEFT,
               command=lambda: self.YokogawaGS200MainMenu()).grid()
        Button(self, padx=25, pady=25, text='LakeShore 336 Temperature Controller', font=(font_size),
               image=lakeshore_img, compound=LEFT,
               command=lambda: self.LakeShore336MainMenu()).grid(column=1, row=2)
        Label(self, text='Automation Menu', font=(font_size)).grid(columnspan=2, column=0)
        Button(self, padx=25, pady=25, text='Automation Menu', font=(font_size),
               command=lambda: self.AutomationMenu()).grid(columnspan=2, column=0)

    def Agilent34410AMainMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text='Connected to:').grid()
        Label(self, text=self.Agilent34410A('ask', '*IDN?')).grid()
        Button(self, text='Configure Device', command=lambda: self.Agilent34410AConfigMenu()).grid()
        Button(self, text='Take a Measurement', command=lambda: self.Agilent34410AMeasurementMenu()).grid()
        Button(self, image=back_img, command=lambda: self.DeviceMen()).grid()

    def Agilent34410AConfigMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text='Display ON', command=lambda: self.Agilent34410A('write', 'DISPlay ON')).grid()
        Button(self, text='Display OFF', command=lambda: self.Agilent34410A('write', 'DISPlay OFF')).grid()
        Button(self, text='Factory Reset Device', command=lambda: self.Agilent34410A('write', '*RST')).grid()
        Button(self, image=back_img, command=lambda: self.Agilent34410AMainMenu()).grid()

    def Agilent34410AMeasurementMenu(self):
        global var
        var = float(var)
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text='Measure Resistance', command=(lambda: self.Agilent34410A('test', 'MEAS?'))).grid()
        Label(self, text=(str((var / .2))) + ' Ohms').grid()
        print (var / .2)
        Button(self, image=back_img, command=lambda: self.Agilent34410AMainMenu()).grid()

    def Agilent34410A(self, option, command):
        settings = open('settings.txt', 'r')
        global var
        adress = settings.readline()
        while adress.rstrip() != 'Agilent34410A':
            adress = settings.readline()
        adress = settings.readline()
        settings.close()
        inst = visa.ResourceManager()
        inst = inst.open_resource(adress.rstrip())
        if option == 'test':
            var = inst.query(command)
            self.Agilent34410AMeasurementMenu()
        if option == 'write':
            inst.write(command)
        if option == 'ask':
            return inst.query(command)
        inst.close()

    def Keithley7002MainMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text='Configure Device', command=lambda: self.Keithley7002ConfigMenu()).grid()
        Button(self, text='Switch Menu', command=lambda: self.Keithley7002SwitchMenu()).grid()
        Button(self, image=back_img, command=lambda: self.DeviceMen()).grid()

    def Keithley7002ConfigMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text='Display ON', command=lambda: self.Keithley7002('write', 'DISPlay:ENABle 1')).grid()
        Button(self, text='Display OFF', command=lambda: self.Keithley7002('write', 'DISPlay:ENABle 0')).grid()
        Button(self, text='Factory Reset Device', command=lambda: self.Keithley7002('write', 'STATus:PRESet')).grid()
        Button(self, image=back_img, command=lambda: self.Keithley7002MainMenu()).grid()

    def Keithley7002SwitchMenu(self):
        card = StringVar()
        inputs = StringVar()
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text='Slot Number (1-10)').grid()
        Entry(self, textvariable=card).grid()
        Label(self, text='Input Number (1-40)').grid()
        Entry(self, textvariable=inputs).grid()
        Button(self, text='Close', command=lambda: self.Keithley7002('write', 'close (@' + str(card.get()) + '!' + str(
            inputs.get()) + ')')).grid()
        Button(self, text='Open All', command=lambda: self.Keithley7002('write', 'open all')).grid()
        Button(self, image=back_img, command=lambda: self.Keithley7002MainMenu()).grid()

    def Keithley7002(self, option, command):
        settings = open('settings.txt', 'r')
        adress = settings.readline()
        while adress.rstrip() != 'Keithley7002':
            adress = settings.readline()
        adress = settings.readline()
        settings.close()
        inst = visa.ResourceManager()
        inst = inst.open_resource(adress.rstrip())
        if option == 'write':
            inst.write(command)
        if option == 'ask':
            return inst.query(command)
        inst.close()

    def YokogawaGS200MainMenu(self):
        ans = StringVar()
        Interval = StringVar()
        SlopeTime = StringVar()
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text='Configure Device').grid()
        Entry(self, textvariable=ans).grid()
        Button(self, text='Send', command=lambda: self.YokogawaGS200('write', ans.get())).grid()
        Label(self, text="Time Interval").grid()
        Entry(self, textvariable=Interval).grid()  # Time Interval (s)
        Button(self, text="Send", command=lambda: self.YokogawaGS200("write", Interval.get())).grid()
        Label(self, text="SlopeTime").grid()
        Entry(self, textvariable=SlopeTime).grid()
        # Button(self, text="Send", command=lambda: self.YokogawaGS200("write", SlopeTIme.get()))
        Button(self, text="Repeat Execution").grid()
        Button(self, text="Pause Execution").grid()
        Button(self, image=back_img, command=lambda: self.DeviceMen()).grid()

    def YokogawaGS200(self, option, command):
        settings = open('settings.txt', 'r')
        adress = settings.readline()
        while adress.rstrip() != 'YokogawaGS200':
            adress = settings.readline()
        adress = settings.readline()
        settings.close()
        inst = visa.ResourceManager()
        inst = inst.open_resource(adress.rstrip())
        if option == 'write':
            inst.write(command)
        if option == 'ask':
            return inst.query(command)
        inst.close()

    def LakeShore336MainMenu(self):
        Kelvin = StringVar()
        TempLim = StringVar()
        High = StringVar()
        Low = StringVar()
        var = 0
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text="Configuration Menu", command=lambda: self.LakeShore336ConfigMenu()).grid()
        Button(self, text="Temperature Readings", command=lambda: self.LakeShore336TempReadMenu()).grid()
        Button(self, text="Heater Settings", command=lambda: self.LakeShore336HeatMenu()).grid()
        Button(self, image=back_img, command=lambda: self.DeviceMen()).grid()

    def LakeShore336ConfigMenu(self):
        var = 0
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text="Power Up Reset Device", command=lambda: self.LakeShore336("write", "*RST")).grid()
        Button(self, text="Factory Reset", command=lambda: self.LakeShore336("write", "DFLT 99")).grid()
        Button(self, text="Brightness Up", command=lambda: self.LakeShore336('write', 'BRIGT 32')).grid()
        Button(self, text="Brightness Down", command=lambda: self.LakeShore336('write', 'BRIGT 0')).grid()
        Button(self, text="Alarm Settings", command=lambda: self.LakeShore336AlarmMenu()).grid()
        Button(self, text="PID Autotune", command=lambda: self.LakeShore336("write", "ATUNE 1,2")).grid()
        Button(self, image=back_img, command=lambda: self.LakeShore336MainMenu()).grid()

    def LakeShore336AlarmMenu(self):
        High = StringVar()
        Low = StringVar()
        var = 0
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text="Alarm High Settings (K)").grid()
        Entry(self, textvariable=High).grid()
        Label(self, text="Alarm Low Settings (K)").grid()
        Entry(self, textvariable=Low).grid()
        Button(self, text="Send", command=lambda: self.LakeShore336("write",
                                                                    "ALARM A,1," + High.get() + "," + Low.get() + ",0,1,1,1")).grid()
        Button(self, text="Alarm Off", command=lambda: self.LakeShore336("write", "ALARM A,0")).grid()
        Button(self, image=back_img, command=lambda: self.LakeShore336MainMenu()).grid()

    def LakeShore336TempReadMenu(self):
        var = 0
        self.destroy()
        Frame.__init__(self)
        self.grid()
        global ans
        global kelv
        Label(self, text=ans).grid()
        Button(self, text="Celsius Reading", command=lambda: self.LakeShore336("celsius", "CRDG? A")).grid()
        Label(self, text=kelv).grid()
        Button(self, text="Kelvin Reading", command=lambda: self.LakeShore336("kelvin", "KRDG? A")).grid()
        Button(self, image=back_img, command=lambda: self.LakeShore336MainMenu()).grid()

    def LakeShore336HeatMenu(self):
        var = 0
        self.destroy()
        Frame.__init__(self)
        self.grid()
        TempLim = StringVar()
        Setpt = StringVar()
        Label(self, text="Temperature Limit (K)").grid()
        Entry(self, textvariable=TempLim).grid()
        Button(self, text="Send", command=lambda: self.LakeShore336("write", "TLIMIT A," + TempLim.get())).grid()
        Label(self, text="Setpoint (K)").grid()
        Entry(self, textvariable=Setpt).grid()
        Button(self, text="Send", command=lambda: self.LakeShore336("write", "SETP 1," + Setpt.get())).grid()
        Label(self, text="Heater Range").grid()
        Button(self, text="High", command=lambda: self.LakeShore336("write", "RANGE 1,3")).grid()
        Button(self, text="Medium", command=lambda: self.LakeShore336("write", "RANGE 1,2")).grid()
        Button(self, text="Low", command=lambda: self.LakeShore336("write", "RANGE 1,1")).grid()
        Button(self, text="OFF", command=lambda: self.LakeShore336("write", "RANGE 1,0")).grid()
        Button(self, image=back_img, command=lambda: self.LakeShore336MainMenu()).grid()

    def LakeShore336(self, option, command):
        global ans
        global kelv
        settings = open('settings.txt', 'r')
        adress = settings.readline()
        while adress.rstrip() != 'LakeShore336':
            adress = settings.readline()
        adress = settings.readline()
        settings.close()
        inst = visa.ResourceManager()
        inst = inst.open_resource(adress.rstrip())
        if option == 'write':
            inst.write(command)
        if option == 'ask':
            return inst.query(command)
        if option == "celsius":
            ans = self.LakeShore336("ask", "CRDG? A")
            self.LakeShore336TempReadMenu()
        if option == "kelvin":
            kelv = self.LakeShore336("ask", "KRDG? A")
            self.LakeShore336TempReadMenu()
        inst.close()

    def AutomationMenu(self):
        global add_process_img
        global recipe
        global recipe_name
        global file_name
        global apply_img
        global exicute_img
        self.destroy()
        Frame.__init__(self)
        self.grid()
        recipe_list = []
        recipe_names_file = open('Process_Recipes.txt', 'r')
        recipe_names = recipe_names_file.readline().rstrip()
        while recipe_names != '':
            recipe_list.append(recipe_names)
            recipe_names = recipe_names_file.readline().rstrip()
        Label(self, text='Select Automation Process', font=(font_size)).grid()
        Button(self, padx=25, pady=25, text='4 Wire Current vs Voltage Resistance Test', font=(font_size),
               command=lambda: self.FourWireCurrentvsVoltaqgeMenu()).grid()
        Button(self, padx=25, pady=25, text='Voltage Vs Current Graph', font=(font_size),
               command=lambda: self.VoltageVsCurrent()).grid()
        Button(self, padx=25, pady=25, text='Temperature Vs Resistance', font=(font_size),
               command=lambda: self.LiveData()).grid()
        Label(self, text='Choose a process recipe', font=(font_size)).grid(row=0, column=1)
        apply(OptionMenu, (self, recipe) + tuple(recipe_list)).grid(row=1, column=1, ipadx=10, ipady=10)
        Button(self, text='', image=apply_img, compound=TOP, command=lambda: self.RecipesMenu('Open', 'Process')).grid(
            row=2, column=1)
        Label(self, text="Save Process As", font=(font_size)).grid(row=0, column=2)
        Entry(self, textvariable=recipe_name, font=(font_size)).grid(row=1, column=2, ipadx=10, ipady=10)
        Button(self, text="Save", image=save_img, font=(font_size), compound=TOP,
               command=lambda: self.RecipesMenu('Save', 'Process')).grid(row=2, column=2)
        Button(self, text='', image=exicute_img, compound=TOP,
               command=lambda: self.UserProgramableTest1Process("UserRecipe")).grid(column=1, row=3)
        Label(self, text='Processes in Que:', font=(font_size)).grid()
        Label(self, text=count, font=(font_size)).grid()
        Button(self, image=back_img, command=lambda: self.DeviceMen()).grid()
        file_name = 'Process_Recipes.txt'

    def LiveData(self):
        global forced
        global count
        global range
        global name
        global measure
        global to
        global fr
        global graph
        global slot
        global outp
        global inp
        global tm
        global add_process_img
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text='Amount forced (ma)').grid()
        Entry(self, textvariable=forced).grid()
        Label(self, text='Input Card Slot Number (1-10)').grid()
        Entry(self, textvariable=slot).grid()
        Label(self, text='Select switch input').grid()
        Entry(self, textvariable=fr).grid()
        Label(self, text='Choose Sensor Input').grid()
        OptionMenu(self, inp, 'A', 'B', 'C', 'D').grid()
        Label(self, text='Name the Excel File that will be created').grid()
        Entry(self, textvariable=name).grid()
        Button(self, image=add_process_img, command=lambda: self.AddProcessToQue()).grid()
        Button(self, image=back_img, command=lambda: self.AutomationMenu()).grid()
        measure = 'Live Data'

    def FourWireCurrentvsVoltaqgeMenu(self):
        global forced
        global count
        global range
        global name
        global measure
        global to
        global fr
        global graph
        global slot
        global recipe
        global recipe_name
        global file_name
        global back_img
        global add_process_img
        global save_img
        global apply_img
        global exicute_img
        global font_size
        recipe_name.set('')
        graph = StringVar()
        graph.set('column')
        root.geometry("650x650")
        self.destroy()
        Frame.__init__(self)
        self.grid()
        recipe_list = []
        recipe_names_file = open('4_Wire_Recipes.txt', 'r')
        recipe_names = recipe_names_file.readline().rstrip()
        while recipe_names != '':
            recipe_list.append(recipe_names)
            recipe_names = recipe_names_file.readline().rstrip()
        Label(self, text='Amount forced (ma)', font=(font_size)).grid(column=0, row=0, rowspan=2)
        Entry(self, textvariable=forced, font=(font_size)).grid(column=0, row=1, ipadx=10, ipady=10)
        Label(self, text='Input Card Slot Number (1-10)', font=(15)).grid(column=0, row=2)
        Entry(self, textvariable=slot, font=(font_size)).grid(column=0, row=3, ipadx=10, ipady=10)
        Label(self, text='Select switch inputs', font=(15)).grid(column=0, row=2, rowspan=4)
        Label(self, text='From:', font=(font_size)).grid(column=0, row=2, rowspan=5)
        Entry(self, textvariable=fr, font=(font_size)).grid(column=0, row=6, ipadx=10, ipady=10)
        Label(self, text='To:', font=(font_size)).grid(column=0, row=7)
        Entry(self, textvariable=to, font=(font_size)).grid(column=0, row=8, ipadx=10, ipady=10)
        Label(self, text='Name of Excel file that will be created:', font=(font_size)).grid(column=0, row=9)
        Entry(self, textvariable=name, font=(font_size)).grid(column=0, row=10, ipadx=30, ipady=10)
        Label(self, text='Pick graph type', font=(font_size)).grid(column=0, row=11)
        OptionMenu(self, graph, 'column').grid(column=0, row=12, ipadx=10,ipady=10)
        Button(self, image=add_process_img, text='Add Process to Que', compound=TOP,
               command=lambda: self.AddProcessToQue()).grid(column=2, row=0)
        Label(self, text='Processes in Que:', font=(font_size)).grid(column=2, row=1)
        Label(self, text=count).grid(column=2, row=2)
        Label(self, text="Choose From Existing Recipe", font=(font_size)).grid(column=1, row=0, rowspan=2)
        apply(OptionMenu, (self, recipe) + tuple(recipe_list)).grid(column=1, row=1, ipadx=10,ipady=10)
        Button(self, text='', image=apply_img, compound=TOP,
               command=lambda: self.RecipesMenu('Open', '4 Wire C vs V')).grid(column=1,
                                                                                                          row=2)
        Label(self, text='New Recipe Name', font=(font_size)).grid(column=1, row=3)
        Entry(self, textvariable=recipe_name, font=(font_size)).grid(column=1, row=4, ipadx=10, ipady=10)
        Button(self, text="Save Recipe", image=save_img, compound=TOP,
               command=lambda: self.RecipesMenu('Save', '4 Wire C vs V')).grid(column=1,
                                                                                                              row=5)
        Button(self, image=back_img, command=lambda: self.AutomationMenu()).grid(column=1, row=6)
        measure = '4 Wire Forced Current vs Voltage'
        file_name = '4_Wire_Recipes.txt'

    def RecipesMenu(self, option, menu):
        global recipe_name
        global file_name
        global forced
        global count
        global range
        global name
        global measure
        global to
        global fr
        global graph
        global inp
        global slot
        global tm
        global count
        if option == 'Save':
            if menu == '4 Wire C vs V':
                recipe_names_file = open(file_name, 'a')
                recipe_names_file.write(str(recipe_name.get()) + '\n')
                recipe_names_file.close()
                new_recipe_file = open(recipe_name.get() + '.txt', 'w')
                new_recipe_file.write(forced.get() + '\n')
                new_recipe_file.write(to.get() + '\n')
                new_recipe_file.write(fr.get() + '\n')
                new_recipe_file.write(graph.get() + '\n')
                new_recipe_file.write(slot.get() + '\n')
                new_recipe_file.write(name.get() + '\n')
                new_recipe_file.close()
                self.FourWireCurrentvsVoltaqgeMenu()
            if menu == 'V Vs C':
                recipe_names_file = open(file_name, 'a')
                recipe_names_file.write(str(recipe_name.get()) + '\n')
                recipe_names_file.close()
                new_recipe_file = open(recipe_name.get() + '.txt', 'w')
                new_recipe_file.write(forced.get() + '\n')
                new_recipe_file.write(to.get() + '\n')
                new_recipe_file.write(fr.get() + '\n')
                new_recipe_file.write(graph.get() + '\n')
                new_recipe_file.write(slot.get() + '\n')
                new_recipe_file.write(name.get() + '\n')
                new_recipe_file.write(inp.get() + '\n')
                new_recipe_file.write(tm.get() + '\n')
                new_recipe_file.close()
                self.VoltageVsCurrent()
            if menu == 'Process':
                recipe_names_file = open(file_name, 'a')
                recipe_names_file.write(str(recipe_name.get()) + '\n')
                recipe_names_file.close()
                new_recipe_file = open(recipe_name.get() + '.txt', 'w')
                process = open('process_que.txt', 'r')
                new_recipe_file.write(str(count).rstrip() + '\n')
                processNumber = 0
                while processNumber < count:
                    measure = process.readline()
                    new_recipe_file.write(str(measure).rstrip() + '\n')
                    forced = process.readline()
                    new_recipe_file.write(str(forced).rstrip() + '\n')
                    range = process.readline()
                    new_recipe_file.write(str(range).rstrip() + '\n')
                    tm = process.readline()
                    new_recipe_file.write(str(tm).rstrip() + '\n')
                    fr = process.readline()
                    new_recipe_file.write(str(fr).rstrip() + '\n')
                    to = process.readline()
                    new_recipe_file.write(str(to).rstrip() + '\n')
                    name = process.readline()
                    new_recipe_file.write(str(name).rstrip() + '\n')
                    graph = process.readline()
                    new_recipe_file.write(str(graph).rstrip() + '\n')
                    rate = process.readline()
                    new_recipe_file.write(str(rate).rstrip() + '\n')
                    wanted_temp = process.readline()
                    new_recipe_file.write(str(wanted_temp).rstrip() + '\n')
                    outp = process.readline()
                    new_recipe_file.write(str(outp).rstrip() + '\n')
                    inp = process.readline()
                    new_recipe_file.write(str(inp).rstrip() + '\n')
                    slot = process.readline()
                    new_recipe_file.write(str(slot).rstrip() + '\n')
                    processNumber += 1
                self.AutomationMenu()
        if option == 'Open':
            if menu == '4 Wire C vs V':
                recipe_file = open(recipe.get() + '.txt', 'r')
                forced.set(recipe_file.readline().rstrip())
                to.set(recipe_file.readline().rstrip())
                fr.set(recipe_file.readline().rstrip())
                graph.set(recipe_file.readline().rstrip())
                slot.set(recipe_file.readline().rstrip())
                name.set(recipe_file.readline().rstrip())
                recipe_file.close()
                self.FourWireCurrentvsVoltaqgeMenu()
            if menu == 'V Vs C':
                recipe_file = open(recipe.get() + '.txt', 'r')
                forced.set(recipe_file.readline().rstrip())
                to.set(recipe_file.readline().rstrip())
                fr.set(recipe_file.readline().rstrip())
                graph.set(recipe_file.readline().rstrip())
                slot.set(recipe_file.readline().rstrip())
                name.set(recipe_file.readline().rstrip())
                inp.set(recipe_file.readline().rstrip())
                tm.set(recipe_file.readline().rstrip())
                recipe_file.close()
                self.VoltageVsCurrent()
            if menu == 'Process':
                recipe_file = open(recipe.get() + '.txt', 'r')
                process = open('process_que.txt', 'w')
                count = int(recipe_file.readline().rstrip())
                processNumber = 0
                print count
                while processNumber < count:
                    measure = recipe_file.readline()
                    process.write(str(measure).rstrip() + '\n')
                    forced = recipe_file.readline()
                    process.write(str(forced).rstrip() + '\n')
                    range = recipe_file.readline()
                    process.write(str(range).rstrip() + '\n')
                    tm = recipe_file.readline()
                    process.write(str(tm).rstrip() + '\n')
                    fr = recipe_file.readline()
                    process.write(str(fr).rstrip() + '\n')
                    to = recipe_file.readline()
                    process.write(str(to).rstrip() + '\n')
                    name = recipe_file.readline()
                    process.write(str(name).rstrip() + '\n')
                    graph = recipe_file.readline()
                    process.write(str(graph).rstrip() + '\n')
                    rate = recipe_file.readline()
                    process.write(str(rate).rstrip() + '\n')
                    wanted_temp = recipe_file.readline()
                    process.write(str(wanted_temp).rstrip() + '\n')
                    outp = recipe_file.readline()
                    process.write(str(outp).rstrip() + '\n')
                    inp = recipe_file.readline()
                    process.write(str(inp).rstrip() + '\n')
                    slot = recipe_file.readline()
                    process.write(str(slot).rstrip() + '\n')
                    processNumber += 1
                self.AutomationMenu()

    def VoltageVsCurrent(self):
        global add_process_img
        self.destroy()
        Frame.__init__(self)
        self.grid()
        global forced
        global count
        global range
        global name
        global measure
        global to
        global fr
        global graph
        global slot
        global outp
        global inp
        global tm
        global recipe
        global recipe_name
        global file_name
        global apply_img
        global exicute_img
        recipe_name.set('')
        recipe_list = []
        root.geometry("650x600")
        recipe_names_file = open('V_Vs_C_Recipes.txt', 'r')
        recipe_names = recipe_names_file.readline().rstrip()
        while recipe_names != '':
            recipe_list.append(recipe_names)
            recipe_names = recipe_names_file.readline().rstrip()
        Label(self, text='Starting Current (ma)', font=(font_size)).grid()
        Entry(self, textvariable=forced, font=(font_size)).grid(ipadx=10, ipady=10)
        Label(self, text='Current Limit (ma)', font=(font_size)).grid()
        Entry(self, textvariable=to, font=(font_size)).grid(ipadx=10, ipady=10)
        Label(self, text='Voltage Limit (Volts)', font=(font_size)).grid()
        Entry(self, textvariable=inp, font=(font_size)).grid(ipadx=10, ipady=10)
        Label(self, text='Current Steps (ma)', font=(font_size)).grid()
        Entry(self, textvariable=tm, font=(font_size)).grid(ipadx=10, ipady=10)
        Label(self, text='Input Card Slot Number (1-10)', font=(font_size)).grid()
        Entry(self, textvariable=slot, font=(font_size)).grid(ipadx=10, ipady=10)
        Label(self, text='Select switch input', font=(font_size)).grid()
        Entry(self, textvariable=fr, font=(font_size)).grid(ipadx=10, ipady=10)
        Label(self, text='Name the Excel file that will be created', font=(font_size)).grid()
        Entry(self, textvariable=name, font=(font_size)).grid(ipadx=10, ipady=10)
        Button(self, image=add_process_img, text='Add Process to Que', font=(font_size), compound=TOP,
               command=lambda: self.AddProcessToQue()).grid(column=2, row=0)
        Label(self, text='Processes in Que:', font=(font_size)).grid(column=2, row=1)
        Label(self, text=count, font=(font_size)).grid(column=2, row=2)
        Label(self, text='New Recipe Name', font=(font_size)).grid(column=1, row=3)
        Entry(self, textvariable=recipe_name, font=(font_size)).grid(column=1, row=4, ipadx=10, ipady=10)
        Button(self, text="Save Recipe", font=(font_size), image=save_img, compound=TOP,
               command=lambda: self.RecipesMenu('Save', 'V Vs C')).grid(column=1,
                                                                                                       row=5)
        Label(self, text="Choose From Existing Recipe", font=(font_size)).grid(column=1, row=0)
        apply(OptionMenu, (self, recipe) + tuple(recipe_list)).grid(column=1, row=1)
        Button(self, text='', image=apply_img, compound=TOP, command=lambda: self.RecipesMenu('Open', 'V Vs C')).grid(
            column=1,
                                                                                                   row=2)
        Button(self, image=back_img, command=lambda: self.AutomationMenu()).grid(column=1, row=6)
        measure = 'VoltageVsCurrent'
        file_name = 'V_Vs_C_Recipes.txt'

    def HeatVsTime(self):
        global add_process_img
        self.destroy()
        Frame.__init__(self)
        self.grid()
        global measure
        global name
        global graph
        global rate
        global count
        global wanted_temp
        global tm
        global outp
        global inp
        Label(self, text='Choose Heater Output').grid()
        OptionMenu(self, outp, '1', '2').grid()
        Label(self, text='Choose Sensor Input').grid()
        OptionMenu(self, inp, 'A', 'B', 'C', 'D').grid()
        Label(self, text='Wanted Temperature (Kelvin)').grid()
        Entry(self, textvariable=wanted_temp).grid()
        Label(self, text='Choose Heating Rate').grid()
        OptionMenu(self, rate, '1', '2', '3').grid()
        Label(self, text='Time interval for checking temperature (Seconds)').grid()
        Entry(self, textvariable=tm).grid()
        Label(self, text='Name Excel file that will be created').grid()
        Entry(self, textvariable=name).grid()
        Button(self, image=add_process_img, command=lambda: self.AddProcessToQue()).grid()
        Button(self, text='Execute Process Que', command=lambda: self.AutoMeasure()).grid()
        Label(self, text='Processes in Que:').grid()
        Label(self, text=count).grid()
        Button(self, image=back_img, command=lambda: self.AutomationMenu()).grid()
        measure = 'Temperature Vs Time'

    def AddProcessToQue(self):
        global tm
        global measure
        global fr
        global to
        global name
        global count
        global forced
        global range
        global graph
        global rate
        global wanted_temp
        global outp
        global inp
        global slot
        process = open('process_que.txt', 'a')
        process.write(str(measure) + '\n')
        process.write(str(forced.get()) + '\n')
        process.write(str(range.get()) + '\n')
        process.write(str(tm.get()) + '\n')
        process.write(str(fr.get()) + '\n')
        process.write(str(to.get()) + '\n')
        process.write(str(name.get()) + '\n')
        process.write(str(graph.get()) + '\n')
        process.write(str(rate.get()) + '\n')
        process.write(str(wanted_temp.get()) + '\n')
        process.write(str(outp.get()) + '\n')
        process.write(str(inp.get()) + '\n')
        process.write(str(slot.get()) + '\n')
        count += 1
        process.close()
        self.AutomationMenu()

    def UserProgramableTest1Process(self, rec):
        global measure
        global tm
        global fr
        global to
        global name
        global worksheet
        global workbook
        global forced
        global range
        global format
        global graph
        global rate
        global wanted_temp
        global inp
        global outp
        global slot
        process = open('process_que.txt', 'r')
        processNumber = 0
        while processNumber < count:
            self.destroy()
            Frame.__init__(self)
            self.grid()
            measure = process.readline()
            forced = process.readline()
            range = process.readline()
            tm = process.readline()
            fr = process.readline()
            to = process.readline()
            name = process.readline()
            graph = process.readline()
            rate = process.readline()
            wanted_temp = process.readline()
            outp = process.readline()
            inp = process.readline()
            slot = process.readline()
            processNumber += 1
            print (name.rstrip() != '')
            if name.rstrip() != '':
                workbook = xlsxwriter.Workbook(str(name).rstrip() + '.xlsx')
                format = workbook.add_format()
                format.set_text_wrap()
                worksheet = workbook.add_worksheet()
            Label(self, text='Currently Running: ' + measure).grid()
            Label(self, text=str((count - processNumber)) + ' More process(s) to go').grid()
            self.AutoMeasure()
        if name.rstrip() != '':
            workbook.close()
        winsound.PlaySound('beep-01.wav', winsound.SND_FILENAME)
        self.AutomationMenu()

    def AutoMeasure(self):
        global name
        global tm
        global measure
        global worksheet
        global workbook
        global fr
        global forced
        global format
        global wanted_temp
        global rate
        global graph
        global to
        global inp
        global outp
        global slot
        global range
        col = 0
        row = 0
        tme = 0
        x = []
        y = []
        i = 0
        if str(measure.rstrip()) == '2 Wire Forced Current vs Voltage':
            self.Keithley7002('write', 'open all')
            self.Keithley7002('write', 'CONF:SLOT' + str(slot).rstrip() + ':POLE 2')
            time.sleep(1)
            if name.rstrip() != '':
                worksheet.write(row, col, 'Current (ma)', format)
                worksheet.write(row, col + 1, 'Voltage (v)', format)
                worksheet.write(row, col + 2, 'Resistance (ohm)', format)
            forced = str(float(forced.rstrip())/1000)
            while int(fr) < int(to) + 1:
                row += 1
                fr = str(fr).rstrip()
                self.Keithley7002('write', 'close (@' + str(slot).rstrip() + '!' + (str(fr)).rstrip() + ')')
                fr = int(fr) + 1
                self.YokogawaGS200('write', 'SENS:REM ON')
                self.YokogawaGS200('write', 'SOUR:FUNC CURR')
                self.YokogawaGS200('write', 'SOUR:RANG ' + forced)
                self.YokogawaGS200('write', 'SOUR:LEV ' + forced)
                self.YokogawaGS200('write', 'OUTP ON')
                time.sleep(.25)
                if name.rstrip() != '':
                    worksheet.write(row, col, '=' + str((float(forced) * 1000)))
                    worksheet.write(row, col + 1, '=' + str(self.Agilent34410A('ask', 'MEAS:VOLT:DC?')))
                    worksheet.write(row, col + 2, '=' + str(
                        (float(self.Agilent34410A('ask', 'MEAS:VOLT:DC?'))) / (float(forced))))
                self.YokogawaGS200('write', 'OUTP OFF')
                self.Keithley7002('write', 'open all')
            if name.rstrip() != '':
                chart = workbook.add_chart({'type': graph.rstrip()})
                chart.add_series({'values': '=Sheet1!$C$2:$C$' + str(row + 1)})
                worksheet.insert_chart('G2', chart)
        if str(measure.rstrip()) == "Live Data":
            matplotlib.pyplot.ion()
            data = 200.00
            forced = str(float(forced.rstrip())/1000)
            self.Keithley7002('write', 'close (@' + str(slot).rstrip() + '!' + (str(fr)).rstrip() + ')')
            self.YokogawaGS200('write', 'SENS:REM ON')
            self.YokogawaGS200('write', 'SOUR:FUNC CURR')
            self.YokogawaGS200('write', 'SOUR:RANG ' + '.1')
            self.YokogawaGS200('write', 'SOUR:LEV ' + forced)
            self.YokogawaGS200('write', 'OUTP ON')
            if name.rstrip() != '':
                worksheet.write(row, col, 'Temperature (K)', format)
                worksheet.write(row, col + 1, 'Resistance (ohms)', format)
            while 1==1:
                data = 0.00
                row += 1
                x.append(float(self.LakeShore336('ask', 'KRDG? ' + inp.rstrip())))
                data = float(self.Agilent34410A('ask', 'MEAS:VOLT:DC?').rstrip())
                y.append(data / float(forced.rstrip()))
                matplotlib.pyplot.plot(x, y)
                matplotlib.pyplot.draw()
                if name.rstrip() != '':
                    worksheet.write(row, col, '=' + self.LakeShore336('ask', 'KRDG? ' + inp.rstrip()))
                    worksheet.write(row, col + 1, '=' + str(data / float(forced)))
            self.YokogawaGS200('write', 'OUTP OFF')
            self.Keithley7002('write', 'open all')
        if str(measure.rstrip()) == 'VoltageVsCurrent':
            voltage = '0'
            x = []
            y = []
            matplotlib.pyplot.ion()
            self.Keithley7002('write', 'close (@' + str(slot).rstrip() + '!' + (str(fr)).rstrip() + ')')
            self.Keithley7002('write', 'CONF:SLOT' + str(slot).rstrip() + ':POLE 2')
            forced = str(float(forced) / 1000)
            to = str(float(to.rstrip())/1000)
            tm = str(float(tm.rstrip())/1000)
            print forced
            print to
            print tm
            if name.rstrip() != '':
                worksheet.write(row, col, 'Current (ma)', format)
                worksheet.write(row, col + 1, 'Voltage (mv)', format)
            while float(inp.rstrip()) >= abs(float(voltage)) and float(to) != float(forced):
                row += 1
                self.YokogawaGS200('write', 'SENS:REM ON')
                self.YokogawaGS200('write', 'SOUR:FUNC CURR')
                self.YokogawaGS200('write', 'SOUR:RANG ' + forced)
                self.YokogawaGS200('write', 'SOUR:LEV ' + forced)
                self.YokogawaGS200('write', 'OUTP ON')
                time.sleep(.75)
                voltage = self.Agilent34410A('ask', 'MEAS:VOLT:DC?').rstrip()
                y.append(float(forced)*1000)
                x.append(float(voltage))
                matplotlib.pyplot.plot(x, y)
                matplotlib.pyplot.draw()
                if name.rstrip() != '':
                    worksheet.write(row, col, '=' + str((float(forced) * 1000)))
                    worksheet.write(row, col + 1, '=' + str(float(voltage)*1000))
                forced = str(float((float(forced)) + (float(tm))))
                print forced
                #time.sleep(.1)
            self.YokogawaGS200('write', 'OUTP OFF')
            self.Keithley7002('write', 'open all')
        if str(measure.rstrip()) == '4 Wire Forced Current vs Voltage':
            self.Keithley7002('write', 'open all')
            self.Keithley7002('write', 'CONF:SLOT' + str(slot).rstrip() + ':POLE 2')
            time.sleep(1)
            if name.rstrip() != '':
                worksheet.write(row, col, 'Current', format)
                worksheet.write(row, col + 1, 'Voltage', format)
                worksheet.write(row, col + 2, 'Resistance', format)
            forced = str(float(forced.rstrip())/1000)
            while int(fr) < int(to) + 1:
                row += 1
                fr = str(fr).rstrip()
                self.Keithley7002('write', 'close (@' + str(slot).rstrip() + '!' + (str(fr)).rstrip() + ')')
                fr = int(fr) + 1
                self.YokogawaGS200('write', 'SENS:REM OFF')
                self.YokogawaGS200('write', 'SOUR:FUNC CURR')
                self.YokogawaGS200('write', 'SOUR:RANG ' + forced)
                self.YokogawaGS200('write', 'SOUR:LEV ' + forced)
                self.YokogawaGS200('write', 'OUTP ON')
                time.sleep(.25)
                if name.rstrip() != '':
                    worksheet.write(row, col, '=' + str(float(forced)*1000))
                    worksheet.write(row, col + 1, '=' + str(self.Agilent34410A('ask', 'MEAS:VOLT:DC?')))
                    worksheet.write(row, col + 2, '=' + str(float(self.Agilent34410A('ask', 'MEAS:VOLT:DC?')) / float(forced)))
                self.YokogawaGS200('write', 'OUTP OFF')
                self.Keithley7002('write', 'open all')
            if name.rstrip() != '':
                chart = workbook.add_chart({'type': graph.rstrip()})
                chart.add_series({'values': '=Sheet1!$B$2:$B$' + str(row + 1)})
                worksheet.insert_chart('A7', chart)
        if str(measure.rstrip()) == 'Temperature Vs Time':
            self.Keithley7002('write', 'open all')
            tm = tm.rstrip()
            wait = tm
            tme = 0.00
            to = self.LakeShore336('ask', 'KRDG? ' + inp.rstrip())
            worksheet.write(row, col, 'Time', format)
            worksheet.write(row, col + 1, 'Temperature', format)
            worksheet.write(row, col + 2, 'Total Time Elapsed (Seconds)', format)
            worksheet.write(row, col + 3, 'Average Kelvins per Second', format)
            worksheet.write(row, col + 4, 'Heating Rate (1-3)', format)
            worksheet.write(row + 1, col + 4, str(rate.rstrip()))
            while float(wanted_temp) > float(to):
                row += 1
                worksheet.write(row, col + 1, '=' + str(self.LakeShore336('ask', 'KRDG? ' + inp.rstrip())))
                worksheet.write(row, col, '=' + str(tme))
                self.LakeShore336('write', 'RANGE ' + outp.rstrip() + ',' + rate.rstrip())
                to = self.LakeShore336('ask', 'KRDG? ' + inp.rstrip())
                time.sleep(float(wait))
                tme += float(tm)
            self.LakeShore336('write', 'RANGE  ' + outp.rstrip() + ',0')
            worksheet.write(1, 2, tme, format)
            worksheet.write(1, 3, '=(B' + str(row) + '-B2)/' + str(tme), format)
            chart = workbook.add_chart({'type': 'scatter'})
            chart.add_series(
                {'categories': '=Sheet1!$A$2:$A$' + str(row + 1), 'values': '=Sheet1!$B$2:$B$' + str(row + 1)})
            worksheet.insert_chart('G2', chart)

root = Tk()
root.title("Measurement System GUI Alpha")
root.geometry("700x600")
app = Application(root)
root.mainloop()
