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

kelv = 0
ans = '0'


class Application(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        self.grid()
        self.DeviceMen()

    def DeviceMen(self):
        root.geometry("300x600")
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
        Label(self, text='Select a device to connect to').grid()
        Button(self, text='Agilent 34410A DMM', command=lambda: self.Agilent34410AMainMenu()).grid()
        Button(self, text='Keithley 7002 Switching Machine', command=lambda: self.Keithley7002MainMenu()).grid()
        Button(self, text='Yokogawa GS200', command=lambda: self.YokogawaGS200MainMenu()).grid()
        Button(self, text='LakeShore 336 Temperature Controller', command=lambda: self.LakeShore336MainMenu()).grid()
        Label(self, text='Automation Menu').grid()
        Button(self, text='Automation Menu', command=lambda: self.AutomationMenu()).grid()

    def Agilent34410AMainMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text='Connected to:').grid()
        Label(self, text=self.Agilent34410A('ask', '*IDN?')).grid()
        Button(self, text='Configure Device', command=lambda: self.Agilent34410AConfigMenu()).grid()
        Button(self, text='Take a Measurement', command=lambda: self.Agilent34410AMeasurementMenu()).grid()
        Button(self, text='Back', command=lambda: self.DeviceMen()).grid()

    def Agilent34410AConfigMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text='Display ON', command=lambda: self.Agilent34410A('write', 'DISPlay ON')).grid()
        Button(self, text='Display OFF', command=lambda: self.Agilent34410A('write', 'DISPlay OFF')).grid()
        Button(self, text='Factory Reset Device', command=lambda: self.Agilent34410A('write', '*RST')).grid()
        Button(self, text='Back', command=lambda: self.Agilent34410AMainMenu()).grid()

    def Agilent34410AMeasurementMenu(self):
        global var
        var = float(var)
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text='Measure Resistance', command=(lambda: self.Agilent34410A('test', 'MEAS?'))).grid()
        Label(self, text=(str((var / .2))) + ' Ohms').grid()
        print (var / .2)
        Button(self, text='Back', command=lambda: self.Agilent34410AMainMenu()).grid()

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
        Button(self, text='Back', command=lambda: self.DeviceMen()).grid()

    def Keithley7002ConfigMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text='Display ON', command=lambda: self.Keithley7002('write', 'DISPlay:ENABle 1')).grid()
        Button(self, text='Display OFF', command=lambda: self.Keithley7002('write', 'DISPlay:ENABle 0')).grid()
        Button(self, text='Factory Reset Device', command=lambda: self.Keithley7002('write', 'STATus:PRESet')).grid()
        Button(self, text='Back', command=lambda: self.Keithley7002MainMenu()).grid()

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
        Button(self, text='Back', command=lambda: self.Keithley7002MainMenu()).grid()

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
        Button(self, text='Back', command=lambda: self.DeviceMen()).grid()

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
        Button(self, text='Back', command=lambda: self.DeviceMen()).grid()

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
        Button(self, text="Back", command=lambda: self.LakeShore336MainMenu()).grid()

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
        Button(self, text="Back", command=lambda: self.LakeShore336MainMenu()).grid()

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
        Button(self, text="Back", command=lambda: self.LakeShore336MainMenu()).grid()

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
        Button(self, text="Back", command=lambda: self.LakeShore336MainMenu()).grid()

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
        root.geometry("300x600")
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text='Select Automation Process').grid()
        Button(self, text='4 Wire Current vs Voltage Resistance Test',
               command=lambda: self.FourWireCurrentvsVoltaqgeMenu()).grid()
        Button(self, text='2 Wire Current vs Voltage Resistance Test',
               command=lambda: self.TwoWireCurrentvsVoltageMenu()).grid()
        Button(self, text='Voltage Vs Current Graph', command=lambda: self.VoltageVsCurrent()).grid()
        Button(self, text='Temperature Vs Resistance', command=lambda: self.LiveData()).grid()
        Button(self, text='Execute Process Que', command=lambda: self.UserProgramableTest1Process("UserRecipe")).grid()
        Label(self, text='Processes in Que:').grid()
        Label(self, text=count).grid()
        Button(self, text="Recipes", command=lambda: self.RecipeMenu()).grid()
        Button(self, text='Back', command=lambda: self.DeviceMen()).grid()

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
        Button(self, text='Add Process to Que', command=lambda: self.AddProcessToQue()).grid()
        Button(self, text='Back', command=lambda: self.AutomationMenu()).grid()
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
        recipe_name.set('')
        graph = StringVar()
        graph.set('column')
        root.geometry("650x400")
        self.destroy()
        Frame.__init__(self)
        self.grid()
        recipe_list = []
        recipe_names_file = open('4_Wire_Recipes.txt', 'r')
        recipe_names = recipe_names_file.readline().rstrip()
        while recipe_names != '':
            recipe_list.append(recipe_names)
            recipe_names = recipe_names_file.readline().rstrip()
        Label(self, text='Amount forced (ma)').grid(column=0, row=0)
        Entry(self, textvariable=forced).grid(column=0, row=1)
        Label(self, text='Input Card Slot Number (1-10)').grid(column=0, row=2)
        Entry(self, textvariable=slot).grid(column=0, row=3)
        Label(self, text='Select switch inputs').grid(column=0, row=4)
        Label(self, text='From:').grid(column=0, row=5)
        Entry(self, textvariable=fr).grid(column=0, row=6)
        Label(self, text='To:').grid(column=0, row=7)
        Entry(self, textvariable=to).grid(column=0, row=8)
        Label(self, text='Name of Excel file that will be created:').grid(column=0, row=9)
        Entry(self, textvariable=name).grid(column=0, row=10)
        Label(self, text='Pick graph type').grid(column=0, row=11)
        OptionMenu(self, graph, 'column').grid(column=0, row=12)
        Button(self, text='Add this Process to Que', command=lambda: self.AddProcessToQue()).grid(column=2, row=0)
        Button(self, text='Execute Process Que', command=lambda: self.UserProgramableTest1Process("UserRecipe")).grid(
            column=2, row=1)
        Label(self, text='Processes in Que:').grid(column=2, row=2)
        Label(self, text=count).grid(column=2, row=3)
        Label(self, text="Choose From Existing Recipe").grid(column=1, row=0)
        apply(OptionMenu, (self, recipe) + tuple(recipe_list)).grid(column=1, row=1)
        Button(self, text='Apply Recipe', command=lambda: self.RecipesMenu('Open', '4 Wire C vs V')).grid(column=1,
                                                                                                          row=2)
        Label(self, text='New Recipe Name').grid(column=1, row=3)
        Entry(self, textvariable=recipe_name).grid(column=1, row=3)
        Button(self, text="Save This Recipe", command=lambda: self.RecipesMenu('Save', '4 Wire C vs V')).grid(column=1,
                                                                                                              row=4)
        Button(self, text='Back', command=lambda: self.AutomationMenu()).grid(column=1, row=15)
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
        global slot
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
    def TwoWireCurrentvsVoltageMenu(self):
        global forced
        global count
        global range
        global name
        global measure
        global to
        global fr
        global slot
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text='Amount forced (ma)').grid()
        Entry(self, textvariable=forced).grid()
        Label(self, text='Input Card Slot Number (1-10)').grid()
        Entry(self, textvariable=slot).grid()
        Label(self, text='Select switch inputs').grid()
        Label(self, text='From:').grid()
        Entry(self, textvariable=fr).grid()
        Label(self, text='To:').grid()
        Entry(self, textvariable=to).grid()
        Label(self, text='Name of Excel file that will be created:').grid()
        Entry(self, textvariable=name).grid()
        Label(self, text='Pick graph type').grid()
        OptionMenu(self, graph, 'column', 'scatter', 'bar').grid()
        Button(self, text='Add this Process to Que', command=lambda: self.AddProcessToQue()).grid()
        Button(self, text='Execute Process Que', command=lambda: self.UserProgramableTest1Process("UserRecipe")).grid()
        Label(self, text='Processes in Que:').grid()
        Label(self, text=count).grid()
        Button(self, text="Choose From Existing Recipe").grid()
        Button(self, text="Save This Recipe").grid()
        Button(self, text='Back', command=lambda: self.AutomationMenu()).grid()
        measure = '2 Wire Forced Current vs Voltage'

    def VoltageVsCurrent(self):
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
        Label(self, text='Starting Current (ma)').grid()
        Entry(self, textvariable=forced).grid()
        Label(self, text='Current Limit (ma)').grid()
        Entry(self, textvariable=to).grid()
        Label(self, text='Voltage Limit (Volts)').grid()
        Entry(self, textvariable=inp).grid()
        Label(self, text='Current Steps (ma)').grid()
        Entry(self, textvariable=tm).grid()
        Label(self, text='Input Card Slot Number (1-10)').grid()
        Entry(self, textvariable=slot).grid()
        Label(self, text='Select switch input').grid()
        Entry(self, textvariable=fr).grid()
        Label(self, text='Name the Excel file that will be created').grid()
        Entry(self, textvariable=name).grid()
        Button(self, text='Add this Process to Que', command=lambda: self.AddProcessToQue()).grid()
        Button(self, text='Execute Process Que', command=lambda: self.UserProgramableTest1Process('UserRecipe')).grid()
        Label(self, text='Processes in Que:').grid()
        Label(self, text=count).grid()
        Button(self, text='Back', command=lambda: self.AutomationMenu()).grid()
        measure = 'VoltageVsCurrent'

    def HeatVsTime(self):
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
        Button(self, text='Add this Process to Que', command=lambda: self.AddProcessToQue()).grid()
        Button(self, text='Execute Process Que', command=lambda: self.AutoMeasure()).grid()
        Label(self, text='Processes in Que:').grid()
        Label(self, text=count).grid()
        Button(self, text='Back', command=lambda: self.AutomationMenu()).grid()
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
                    worksheet.write(row, col + 1, '=' + str(self.YokogawaGS200('ask', 'MEAS?')))
                    worksheet.write(row, col + 2, '=' + str(
                        float(self.YokogawaGS200('ask', 'MEAS?')) / (float(forced))))
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
            #self.YokogawaGS200('write', 'OUTP OFF')
            #self.Keithley7002('write', 'open all')
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

    def RecipeMenu (self, option, type):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        if option == 'Open':
            if type == '2 Wire V vs C':
                recipe_names = open('2_Wire_Recipes.txt','r')
                name = recipe_names.readline().rstrip()
                while name != '':
                    Button(self, text=name).grid()
                    name= recipe_names.readline().rstrip()
        Button(self, text="Back", command=lambda: self.AutomationMenu()).grid()

    def MakeRecipe (self):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text="Four Wire Measurement", command=lambda: self.RecipeFourWire()).grid()
        Button(self, text="Two Wire Measurement", command=lambda: self.RecipeTwoWire()).grid()
        Button(self, text="Current Vs. Voltage").grid()
        Button(self, text="Temperature Vs. Resistance").grid()
        Button(self, text="Back", command=lambda: self.RecipeMenu()).grid()

    def RecipeFourWire (self):
        global forced
        global count
        global range
        global name
        global measure
        global to
        global fr
        global graph
        global slot
        global wire
        graph = StringVar()
        graph.set('column')
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text='Amount forced (Amps)').grid()
        Entry(self, textvariable=forced).grid()
        Label(self, text='Range (Amps):').grid()
        Entry(self, textvariable=range).grid()
        Label(self, text='Input Card Slot Cumner (1-10)').grid()
        Entry(self, textvariable=slot).grid()
        Label(self, text='Select switch inputs').grid()
        Label(self, text='From:').grid()
        Entry(self, textvariable=fr).grid()
        Label(self, text='To:').grid()
        Entry(self, textvariable=to).grid()
        Label(self, text='Name of Excel file that will be created:').grid()
        Entry(self, textvariable=name).grid()
        Label(self, text='Pick graph type').grid()
        OptionMenu(self, graph, 'column', 'scatter', 'bar').grid()
        Label(self, text="Save This Recipe As:").grid()
        Entry(self, textvarible=wire).grid()
        Button(self, text="Save Recipe").grid()
        Button(self, text="Back", command=lambda: self.MakeRecipe()).grid()
    def RecipeTwoWire (self):
        global forced
        global count
        global range
        global name
        global measure
        global to
        global fr
        global slot
        global name
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Label(self, text='Amount forced (Amps)').grid()
        Entry(self, textvariable=forced).grid()
        Label(self, text='Range (Amps):').grid()
        Entry(self, textvariable=range).grid()
        Label(self, text='Input Card Slot Cumner (1-10)').grid()
        Entry(self, textvariable=slot).grid()
        Label(self, text='Select switch inputs').grid()
        Label(self, text='From:').grid()
        Entry(self, textvariable=fr).grid()
        Label(self, text='To:').grid()
        Entry(self, textvariable=to).grid()
        Label(self, text='Name of Excel file that will be created:').grid()
        Entry(self, textvariable=name).grid()
        Label(self, text='Pick graph type').grid()
        OptionMenu(self, graph, 'column', 'scatter', 'bar').grid()
        Label(self, text="Save Recipe Name As:").grid()
        Entry(self, textvariable=name).grid()
        Button(self, text="Save Recipe").grid()
        Button(self, text="Back", command=lambda: self.MakeRecipe()).grid()

    def ExistingRecipes (self):
        self.destroy()
        Frame.__init__(self)
        self.grid()
        Button(self, text="Back", command=lambda: self.RecipeMenu()).grid()


root = Tk()
root.title("Measurement System GUI Alpha")
root.geometry("700x600")
app = Application(root)
root.mainloop()
