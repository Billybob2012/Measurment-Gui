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
import numpy

kelv = 0
ans = '0'


class Application(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        self.pack()
        self.DeviceMen()

    def DeviceMen(self):
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
        self.pack()
        Label(self, text='Select a device to connect to').pack()
        Button(self, text='Agilent 34410A DMM', command=lambda: self.Agilent34410AMainMenu()).pack()
        Button(self, text='Keithley 7002 Switching Machine', command=lambda: self.Keithley7002MainMenu()).pack()
        Button(self, text='Yokogawa GS200', command=lambda: self.YokogawaGS200MainMenu()).pack()
        Button(self, text='LakeShore 336 Temperature Controller', command=lambda: self.LakeShore336MainMenu()).pack()
        Label(self, text='Automation Menu').pack()
        Button(self, text='Automation Menu', command=lambda: self.AutomationMenu()).pack()
        Button(self, text='Temperature Vs Resistance Graph', command=lambda: self.LiveData()).pack()

    def Agilent34410AMainMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self, text='Connected to:').pack()
        Label(self, text=self.Agilent34410A('ask', '*IDN?')).pack()
        Button(self, text='Configure Device', command=lambda: self.Agilent34410AConfigMenu()).pack()
        Button(self, text='Take a Measurement', command=lambda: self.Agilent34410AMeasurementMenu()).pack()
        Button(self, text='Back', command=lambda: self.DeviceMen()).pack()

    def Agilent34410AConfigMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button(self, text='Display ON', command=lambda: self.Agilent34410A('write', 'DISPlay ON')).pack()
        Button(self, text='Display OFF', command=lambda: self.Agilent34410A('write', 'DISPlay OFF')).pack()
        Button(self, text='Factory Reset Device', command=lambda: self.Agilent34410A('write', '*RST')).pack()
        Button(self, text='Back', command=lambda: self.Agilent34410AMainMenu()).pack()

    def Agilent34410AMeasurementMenu(self):
        global var
        var = float(var)
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button(self, text='Measure Resistance', command=(lambda: self.Agilent34410A('test', 'MEAS?'))).pack()
        Label(self, text=(str((var / .2))) + ' Ohms').pack()
        print (var / .2)
        Button(self, text='Back', command=lambda: self.Agilent34410AMainMenu()).pack()

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
        self.pack()
        Button(self, text='Configure Device', command=lambda: self.Keithley7002ConfigMenu()).pack()
        Button(self, text='Switch Menu', command=lambda: self.Keithley7002SwitchMenu()).pack()
        Button(self, text='Back', command=lambda: self.DeviceMen()).pack()

    def Keithley7002ConfigMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button(self, text='Display ON', command=lambda: self.Keithley7002('write', 'DISPlay:ENABle 1')).pack()
        Button(self, text='Display OFF', command=lambda: self.Keithley7002('write', 'DISPlay:ENABle 0')).pack()
        Button(self, text='Factory Reset Device', command=lambda: self.Keithley7002('write', 'STATus:PRESet')).pack()
        Button(self, text='Back', command=lambda: self.Keithley7002MainMenu()).pack()

    def Keithley7002SwitchMenu(self):
        card = StringVar()
        inputs = StringVar()
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self, text='Slot Number (1-10)').pack()
        Entry(self, textvariable=card).pack()
        Label(self, text='Input Number (1-40)').pack()
        Entry(self, textvariable=inputs).pack()
        Button(self, text='Close', command=lambda: self.Keithley7002('write', 'close (@' + str(card.get()) + '!' + str(
            inputs.get()) + ')')).pack()
        Button(self, text='Open All', command=lambda: self.Keithley7002('write', 'open all')).pack()
        Button(self, text='Back', command=lambda: self.Keithley7002MainMenu()).pack()

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
        self.pack()
        Button(self, text='Configure Device').pack()
        Entry(self, textvariable=ans).pack()
        Button(self, text='Send', command=lambda: self.YokogawaGS200('write', ans.get())).pack()
        Label(self, text="Time Interval").pack()
        Entry(self, textvariable=Interval).pack()  # Time Interval (s)
        Button(self, text="Send", command=lambda: self.YokogawaGS200("write", Interval.get())).pack()
        Label(self, text="SlopeTime").pack()
        Entry(self, textvariable=SlopeTime).pack()
        Button(self, text="Send", command=lambda: self.YokogawaGS200("write", SlopeTIme.get()))
        Button(self, text="Repeat Execution").pack()
        Button(self, text="Pause Execution").pack()
        Button(self, text='Back', command=lambda: self.DeviceMen()).pack()

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
        self.pack()
        Button(self, text="Configuration Menu", command=lambda: self.LakeShore336ConfigMenu()).pack()
        Button(self, text="Temperature Readings", command=lambda: self.LakeShore336TempReadMenu()).pack()
        Button(self, text="Heater Settings", command=lambda: self.LakeShore336HeatMenu()).pack()
        Button(self, text='Back', command=lambda: self.DeviceMen()).pack()

    def LakeShore336ConfigMenu(self):
        var = 0
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button(self, text="Power Up Reset Device", command=lambda: self.LakeShore336("write", "*RST")).pack()
        Button(self, text="Factory Reset", command=lambda: self.LakeShore336("write", "DFLT 99")).pack()
        Button(self, text="Brightness Up", command=lambda: self.LakeShore336('write', 'BRIGT 32')).pack()
        Button(self, text="Brightness Down", command=lambda: self.LakeShore336('write', 'BRIGT 0')).pack()
        Button(self, text="Alarm Settings", command=lambda: self.LakeShore336AlarmMenu()).pack()
        Button(self, text="PID Autotune", command=lambda: self.LakeShore336("write", "ATUNE 1,2")).pack()
        Button(self, text="Back", command=lambda: self.LakeShore336MainMenu()).pack()

    def LakeShore336AlarmMenu(self):
        High = StringVar()
        Low = StringVar()
        var = 0
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self, text="Alarm High Settings (K)").pack()
        Entry(self, textvariable=High).pack()
        Label(self, text="Alarm Low Settings (K)").pack()
        Entry(self, textvariable=Low).pack()
        Button(self, text="Send", command=lambda: self.LakeShore336("write",
                                                                    "ALARM A,1," + High.get() + "," + Low.get() + ",0,1,1,1")).pack()
        Button(self, text="Alarm Off", command=lambda: self.LakeShore336("write", "ALARM A,0")).pack()
        Button(self, text="Back", command=lambda: self.LakeShore336MainMenu()).pack()

    def LakeShore336TempReadMenu(self):
        var = 0
        self.destroy()
        Frame.__init__(self)
        self.pack()
        global ans
        global kelv
        Label(self, text=ans).pack()
        Button(self, text="Celsius Reading", command=lambda: self.LakeShore336("celsius", "CRDG? A")).pack()
        Label(self, text=kelv).pack()
        Button(self, text="Kelvin Reading", command=lambda: self.LakeShore336("kelvin", "KRDG? A")).pack()
        Button(self, text="Back", command=lambda: self.LakeShore336MainMenu()).pack()

    def LakeShore336HeatMenu(self):
        var = 0
        self.destroy()
        Frame.__init__(self)
        self.pack()
        TempLim = StringVar()
        Setpt = StringVar()
        Label(self, text="Temperature Limit (K)").pack()
        Entry(self, textvariable=TempLim).pack()
        Button(self, text="Send", command=lambda: self.LakeShore336("write", "TLIMIT A," + TempLim.get())).pack()
        Label(self, text="Setpoint (K)").pack()
        Entry(self, textvariable=Setpt).pack()
        Button(self, text="Send", command=lambda: self.LakeShore336("write", "SETP 1," + Setpt.get())).pack()
        Label(self, text="Heater Range").pack()
        Button(self, text="High", command=lambda: self.LakeShore336("write", "RANGE 1,3")).pack()
        Button(self, text="Medium", command=lambda: self.LakeShore336("write", "RANGE 1,2")).pack()
        Button(self, text="Low", command=lambda: self.LakeShore336("write", "RANGE 1,1")).pack()
        Button(self, text="OFF", command=lambda: self.LakeShore336("write", "RANGE 1,0")).pack()
        Button(self, text="Back", command=lambda: self.LakeShore336MainMenu()).pack()

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
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self, text='Select Automation Process').pack()
        Button(self, text='4 Wire Current vs Voltage Resistance Test',
               command=lambda: self.FourWireCurrentvsVoltaqgeMenu()).pack()
        Button(self, text='2 Wire Current vs Voltage Resistance Test',
               command=lambda: self.TwoWireCurrentvsVoltageMenu()).pack()
        Button(self, text='Heat Vs Time', command=lambda: self.HeatVsTime()).pack()
        Button(self, text='Voltage Vs Current Graph', command=lambda: self.VoltageVsCurrent()).pack()
        Button(self, text='Execute Process Que', command=lambda: self.UserProgramableTest1Process()).pack()
        Label(self, text='Processes in Que:').pack()
        Label(self, text=count).pack()
        Button(self, text='Back', command=lambda: self.DeviceMen()).pack()

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
        self.pack()
        Label(self, text='Amount forced (ma)').pack()
        Entry(self, textvariable=forced).pack()
        Label(self, text='Input Card Slot Number (1-10)').pack()
        Entry(self, textvariable=slot).pack()
        Label(self, text='Select switch input').pack()
        Entry(self, textvariable=fr).pack()
        Label(self, text='Choose Sensor Input').pack()
        OptionMenu(self, inp, 'A', 'B', 'C', 'D').pack()
        Button(self, text='Execute', command=lambda: self.AutoMeasure()).pack()
        Button(self, text='Back', command=lambda: self.DeviceMen()).pack()
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
        graph = StringVar()
        graph.set('column')
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self, text='Amount forced (ma)').pack()
        Entry(self, textvariable=forced).pack()
        Label(self, text='Input Card Slot Number (1-10)').pack()
        Entry(self, textvariable=slot).pack()
        Label(self, text='Select switch inputs').pack()
        Label(self, text='From:').pack()
        Entry(self, textvariable=fr).pack()
        Label(self, text='To:').pack()
        Entry(self, textvariable=to).pack()
        Label(self, text='Name of Excel file that will be created:').pack()
        Entry(self, textvariable=name).pack()
        Label(self, text='Pick graph type').pack()
        OptionMenu(self, graph, 'column', 'scatter', 'bar').pack()
        Button(self, text='Add this Process to Que', command=lambda: self.AddProcessToQue()).pack()
        Button(self, text='Execute Process Que', command=lambda: self.UserProgramableTest1Process()).pack()
        Label(self, text='Processes in Que:').pack()
        Label(self, text=count).pack()
        Button(self, text='Back', command=lambda: self.AutomationMenu()).pack()
        measure = '4 Wire Forced Current vs Voltage'

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
        self.pack()
        Label(self, text='Amount forced (ma)').pack()
        Entry(self, textvariable=forced).pack()
        Label(self, text='Input Card Slot Number (1-10)').pack()
        Entry(self, textvariable=slot).pack()
        Label(self, text='Select switch inputs').pack()
        Label(self, text='From:').pack()
        Entry(self, textvariable=fr).pack()
        Label(self, text='To:').pack()
        Entry(self, textvariable=to).pack()
        Label(self, text='Name of Excel file that will be created:').pack()
        Entry(self, textvariable=name).pack()
        Label(self, text='Pick graph type').pack()
        OptionMenu(self, graph, 'column', 'scatter', 'bar').pack()
        Button(self, text='Add this Process to Que', command=lambda: self.AddProcessToQue()).pack()
        Button(self, text='Execute Process Que', command=lambda: self.UserProgramableTest1Process()).pack()
        Label(self, text='Processes in Que:').pack()
        Label(self, text=count).pack()
        Button(self, text='Back', command=lambda: self.AutomationMenu()).pack()
        measure = '2 Wire Forced Current vs Voltage'

    def VoltageVsCurrent(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
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
        Label(self, text='Starting Current (ma)').pack()
        Entry(self, textvariable=forced).pack()
        Label(self, text='Current Limit (ma)').pack()
        Entry(self, textvariable=to).pack()
        Label(self, text='Voltage Limit (Volts)').pack()
        Entry(self, textvariable=name).pack()
        Label(self, text='Current Steps (ma)').pack()
        Entry(self, textvariable=tm).pack()
        Label(self, text='Input Card Slot Number (1-10)').pack()
        Entry(self, textvariable=slot).pack()
        Label(self, text='Select switch input').pack()
        Entry(self, textvariable=fr).pack()
        Button(self, text='Add this Process to Que', command=lambda: self.AddProcessToQue()).pack()
        Button(self, text='Execute Process Que', command=lambda: self.UserProgramableTest1Process()).pack()
        Label(self, text='Processes in Que:').pack()
        Label(self, text=count).pack()
        Button(self, text='Back', command=lambda: self.AutomationMenu()).pack()
        measure = 'VoltageVsCurrent'

    def HeatVsTime(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
        global measure
        global name
        global graph
        global rate
        global count
        global wanted_temp
        global tm
        global outp
        global inp
        Label(self, text='Choose Heater Output').pack()
        OptionMenu(self, outp, '1', '2').pack()
        Label(self, text='Choose Sensor Input').pack()
        OptionMenu(self, inp, 'A', 'B', 'C', 'D').pack()
        Label(self, text='Wanted Temperature (Kelvin)').pack()
        Entry(self, textvariable=wanted_temp).pack()
        Label(self, text='Choose Heating Rate').pack()
        OptionMenu(self, rate, '1', '2', '3').pack()
        Label(self, text='Time interval for checking temperature (Seconds)').pack()
        Entry(self, textvariable=tm).pack()
        Label(self, text='Name Excel file that will be created').pack()
        Entry(self, textvariable=name).pack()
        Button(self, text='Add this Process to Que', command=lambda: self.AddProcessToQue()).pack()
        Button(self, text='Execute Process Que', command=lambda: self.AutoMeasure()).pack()
        Label(self, text='Processes in Que:').pack()
        Label(self, text=count).pack()
        Button(self, text='Back', command=lambda: self.AutomationMenu()).pack()
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

    def UserProgramableTest1Process(self):
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
        processNumber = 0
        process = open('process_que.txt', 'r')
        while processNumber < count:
            self.destroy()
            Frame.__init__(self)
            self.pack()
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
            processNumber = processNumber + 1
            workbook = xlsxwriter.Workbook(str(name).rstrip() + '.xlsx')
            format = workbook.add_format()
            format.set_text_wrap()
            worksheet = workbook.add_worksheet()
            Label(self, text='Currently Running: ' + measure).pack()
            Label(self, text=str((count - processNumber)) + ' More process(s) to go').pack()
            self.AutoMeasure()
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
        forced = str(float(forced.rstrip())/1000)
        if str(measure.rstrip()) == '2 Wire Forced Current vs Voltage':
            self.Keithley7002('write', 'open all')
            self.Keithley7002('write', 'CONF:SLOT' + str(slot).rstrip() + ':POLE 2')
            time.sleep(1)
            worksheet.write(row, col, 'Current', format)
            worksheet.write(row, col + 1, 'Voltage', format)
            worksheet.write(row, col + 2, 'Resistance', format)
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
                worksheet.write(row, col, '=' + forced)
                worksheet.write(row, col + 1, '=' + str(self.YokogawaGS200('ask', 'MEAS?')))
                worksheet.write(row, col + 2, '=' + str(
                    float(self.YokogawaGS200('ask', 'MEAS?')) / float(str(forced.rstrip()))))
                self.YokogawaGS200('write', 'OUTP OFF')
                self.Keithley7002('write', 'open all')
            chart = workbook.add_chart({'type': graph.rstrip()})
            chart.add_series({'values': '=Sheet1!$C$2:$C$' + str(row + 1)})
            worksheet.insert_chart('G2', chart)
        if str(measure.rstrip()) == "Live Data":
            tm = tm.get()
            fr = fr.get()
            inp = inp.get()
            slot = slot.get()
            forced = forced.get()
            matplotlib.pyplot.ion()
            data = 200.00
            forced = str(float(forced.rstrip())/1000)
            self.Keithley7002('write', 'close (@' + str(slot).rstrip() + '!' + (str(fr)).rstrip() + ')')
            self.YokogawaGS200('write', 'SENS:REM ON')
            self.YokogawaGS200('write', 'SOUR:FUNC CURR')
            self.YokogawaGS200('write', 'SOUR:RANG ' + '.1')
            self.YokogawaGS200('write', 'SOUR:LEV ' + forced)
            self.YokogawaGS200('write', 'OUTP ON')
            while data / (float(forced)) >= 10:
                data = 0.00
                x.append(float(self.LakeShore336('ask', 'KRDG? ' + inp.rstrip())))
                data = float(self.Agilent34410A('ask', 'MEAS:VOLT:DC?').rstrip())
                y.append(data / float(forced.rstrip()))
                matplotlib.pyplot.plot(x, y)
                matplotlib.pyplot.draw()
            self.YokogawaGS200('write', 'OUTP OFF')
            self.Keithley7002('write', 'open all')
        if str(measure.rstrip()) == 'VoltageVsCurrent':
            voltage = '0'
            x = []
            y = []
            matplotlib.pyplot.ion()
            self.Keithley7002('write', 'close (@' + str(slot).rstrip() + '!' + (str(fr)).rstrip() + ')')
            self.Keithley7002('write', 'CONF:SLOT' + str(slot).rstrip() + ':POLE 2')
            print (float(name.rstrip()) >= float(voltage.rstrip()))
            print (float(to.rstrip()) / 1000 >= float(forced.rstrip()) / 1000)
            forced = str((float(forced) / 1000))
            to = str(float(to.rstrip())/1000)
            tm = str(float(tm.rstrip())/1000)
            print forced
            print to
            print tm
            while float(name.rstrip()) >= float(voltage.rstrip()) and float(to) >= float(forced):
                print 'Entered While Loop'
                self.YokogawaGS200('write', 'SENS:REM ON')
                self.YokogawaGS200('write', 'SOUR:FUNC CURR')
                self.YokogawaGS200('write', 'SOUR:RANG ' + forced)
                self.YokogawaGS200('write', 'SOUR:LEV ' + forced)
                self.YokogawaGS200('write', 'OUTP ON')
                voltage = self.Agilent34410A('ask', 'MEAS:VOLT:DC?').rstrip()
                x.append(float(forced.rstrip()))
                y.append(float(voltage.rstrip()))
                matplotlib.pyplot.plot(x, y)
                matplotlib.pyplot.draw()
                forced = str((float(forced)) + (float(tm)))
                print float(forced.rstrip())
            self.YokogawaGS200('write', 'OUTP OFF')
            self.Keithley7002('write', 'open all')
        if str(measure.rstrip()) == '4 Wire Forced Current vs Voltage':
            self.Keithley7002('write', 'open all')
            self.Keithley7002('write', 'CONF:SLOT' + str(slot).rstrip() + ':POLE 2')
            time.sleep(1)
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
                worksheet.write(row, col, '=' + forced)
                worksheet.write(row, col + 1, '=' + str(self.Agilent34410A('ask', 'MEAS:VOLT:DC?')))
                worksheet.write(row, col + 2, '=' + str(
                    float(self.Agilent34410A('ask', 'MEAS:VOLT:DC?')) / float(str(forced.rstrip()))))
                self.YokogawaGS200('write', 'OUTP OFF')
                self.Keithley7002('write', 'open all')
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
root.geometry("600x500")
app = Application(root)
root.mainloop()
