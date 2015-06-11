#Megha
try:
    import visa
except ImportError:
    print 'Please install all libraries'
from Tkinter import *
import time
try:
    import xlsxwriter
except:
    print 'Please install all libraries'
from collections import Counter
import serial
ans='0'
kelv = "0"
class Application(Frame):
    def __init__(self, master):
        Frame.__init__(self,master)
        self.pack()
        self.DeviceMen()
    def DeviceMen(self):
        global count
        global var
        var = '0'
        count = 0
        process = open('process_que.txt' , 'w') #Overrides the old process Que document with a blank one on startup
        process.close()  #closes the file
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self,text='Select a device to connect to').pack()
        Button(self,text='Agilent 34410A DMM',command=lambda:self.Agilent34410AMainMenu()).pack()
        Button(self,text='Keithley 7002 Switching Machine',command=lambda:self.Keithley7002MainMenu()).pack()
        Button(self,text='Yokogawa GS200',command=lambda:self.YokogawaGS200MainMenu()).pack()
        Button(self,text='LakeShore 336 Temperature Controller',command=lambda:self.LakeShore336MainMenu()).pack()
        Button(self,text='Arduino Board',command=lambda:self.ArduinoMenu()).pack()
        Label(self,text='Automation Menu').pack()
        Button(self,text='Automation Menu',command=lambda:self.AutomationMenu()).pack()
    def Agilent34410AMainMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self,text='Connected to:').pack()
        Label(self,text=self.Agilent34410A('ask','*IDN?')).pack()
        Button(self,text='Configure Device',command=lambda:self.Agilent34410AConfigMenu()).pack()
        Button(self,text='Take a Measurement',command=lambda:self.Agilent34410AMeasurementMenu()).pack()
        Button(self,text='Back',command=lambda:self.DeviceMen()).pack()
    def Agilent34410AConfigMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button(self,text='Display ON',command=lambda:self.Agilent34410A('write','DISPlay ON')).pack()
        Button(self,text='Display OFF',command=lambda:self.Agilent34410A('write','DISPlay OFF')).pack()
        Button(self,text='Factory Reset Device',command=lambda:self.Agilent34410A('write','*RST')).pack()
        Button(self,text='Back',command=lambda:self.Agilent34410AMainMenu()).pack()
    def Agilent34410AMeasurementMenu(self):
        global var
        var = float(var)
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button(self,text='Measure Resistance',command=(lambda:self.Agilent34410A('test','MEAS?'))).pack()
        Label(self,text=(str((var/.2)))+' Ohms').pack()
        print (var/.2)
        Button(self,text='Back',command=lambda:self.Agilent34410AMainMenu()).pack()
    def Agilent34410A(self, option, command):
        settings = open('settings.txt' , 'r')
        global var
        adress = settings.readline()
        while adress.rstrip() !='Agilent34410A':
            adress = settings.readline()
        adress = settings.readline()
        settings.close()
        inst = visa.ResourceManager()
        inst = inst.open_resource(adress.rstrip())
        if option == 'test':
            var=inst.query(command)
            self.Agilent34410AMeasurementMenu()
        if option =='write':
            inst.write(command)
        if option == 'ask':
            return inst.query(command)
        inst.close()
    def Keithley7002MainMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button(self,text='Configure Device', command=lambda: self.Keithley7002ConfigMenu()).pack()
        Button(self, text = 'Switch Menu',command=lambda:self.Keithley7002SwitchMenu()).pack()
        Button(self,text='Back',command=lambda:self.DeviceMen()).pack()
    def Keithley7002ConfigMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button(self,text='Display ON',command=lambda:self.Keithley7002('write','DISPlay:ENABle 1')).pack()
        Button(self,text='Display OFF',command=lambda:self.Keithley7002('write','DISPlay:ENABle 0')).pack()
        Button(self,text='Factory Reset Device',command=lambda:self.Keithley7002('write','STATus:PRESet')).pack()
        Button(self,text='Back',command=lambda:self.Keithley7002MainMenu()).pack()
    def Keithley7002SwitchMenu(self):
        card = StringVar()
        inputs = StringVar()
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self,text='Slot Number (1-10)').pack()
        Entry(self,textvariable=card).pack()
        Label(self,text='Input Number (1-40)').pack()        
        Entry(self,textvariable=inputs).pack()
        Button(self,text='Close',command=lambda:self.Keithley7002('write','close (@'+str(card.get())+'!'+str(inputs.get())+')')).pack()
        Button(self,text='Open All',command=lambda:self.Keithley7002('write','open all')).pack()
        Button(self,text='Back',command=lambda:self.Keithley7002MainMenu()).pack()
    def Keithley7002(self, option, command):
        settings = open('settings.txt' , 'r')
        adress = settings.readline()
        while adress.rstrip() !='Keithley7002':
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
        Button(self,text='Configure Device').pack()
        Entry(self,textvariable=ans).pack()
        Button(self,text='Send',command=lambda:self.YokogawaGS200('write',ans.get())).pack()
        Label (self, text = "Time Interval").pack()
        Entry (self, textvariable = Interval).pack() #Time Interval (s)
        Button (self, text = "Send", command= lambda:self.YokogawaGS200("write", Interval.get())).pack()
        Label (self, text = "SlopeTime").pack()
        Entry (self, textvariable = SlopeTime).pack()
        Button (self, text = "Send", command = lambda: self. YokogawaGS200("write", SlopeTIme.get()))
        Button (self, text = "Repeat Execution").pack()
        Button (self, text = "Pause Execution").pack()
        Button(self,text='Back',command=lambda:self.DeviceMen()).pack()
    def YokogawaGS200(self, option, command):
        settings = open('settings.txt' , 'r')
        adress = settings.readline()
        while adress.rstrip() !='YokogawaGS200':
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
    def ArduinoMenu(self):
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button(self,text='Raise Device',command=lambda:self.ArduinoBoard('write','1')).pack()
        Button(self,text='Lower Device',command=lambda:self.ArduinoBoard('write','2')).pack()
        Button(self,text='Back',command=lambda:self.DeviceMen()).pack()
    def ArduinoBoard(self,option,command):
        settings=open('settings.txt' , 'r')
        adress=settings.readline()
        while adress.rstrip()!='Arduino Board':
            adress=settings.readline()
        adress=settings.readline()
        settings.close()
        try:
            arduino=serial.Serial(adress.rstrip(),9600)
        except serial.SerialException:
            print 'No Arduino Baord Found on '+adress.rstrip()
        # arduino=serial.Serial(adress.rstrip(),9600)
        time.sleep(.5)
        if option=='write':
            try:
                arduino.write(command)
            except:
                print ''
        if option=='ask':
            print 'nothing to do'
    def LakeShore336MainMenu(self):
        Kelvin=StringVar ()
        TempLim = StringVar()
        High = StringVar()
        Low = StringVar()
        var=0
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button (self, text = "Configuration Menu", command = lambda:self.LakeShore336ConfigMenu()).pack()
        Button (self, text = "Temperature Readings", command = lambda:self.LakeShore336TempReadMenu()).pack()
        Button (self, text = "Heater Settings", command = lambda:self.LakeShore336HeatMenu()).pack()
        Button (self,text='Back',command=lambda:self.DeviceMen()).pack()
    def LakeShore336ConfigMenu(self):
        var=0
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Button (self, text = "Reset Device", command = lambda:self.LakeShore336("write", "*RST")).pack()
        Button(self, text = "Brightness Up",command=lambda:self.LakeShore336('write','BRIGT 32')).pack()
        Button(self, text = "Brightness Down",command=lambda:self.LakeShore336('write','BRIGT 0')).pack()
        Button (self, text = "Alarm Settings", command=lambda:self.LakeShore336AlarmMenu()).pack()
        Button (self, text = "Back",command = lambda:self.LakeShore336MainMenu()).pack()
    def LakeShore336AlarmMenu(self):
        High = StringVar()
        Low = StringVar()
        var=0
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label (self, text = "Alarm High Settings (K)").pack()
        Entry (self, textvariable = High).pack()
        Label (self, text = "Alarm Low Settings (K)").pack()
        Entry (self, textvariable = Low).pack()
        Button (self, text = "Send", command = lambda:self.LakeShore336("write", "ALARM A,1," + High.get() +"," + Low.get() + ",0,1,1,1")).pack()
        Button (self, text= "Alarm Off", command = lambda:self.LakeShore336("write", "ALARM A,0")).pack()
        Button (self, text = "Back", command = lambda:self.LakeShore336MainMenu()).pack()
    def LakeShore336TempReadMenu(self):
        var=0
        self.destroy()
        Frame.__init__(self)
        self.pack()
        global ans
        global kelv
        Label (self, text = ans).pack()
        Button (self, text = "Celsius Reading", command = lambda:self.LakeShore336("celsius", "CRDG? A")).pack()
        Label (self, text = kelv).pack()
        Button (self, text = "Kelvin Reading", command = lambda:self.LakeShore336("kelvin", "KRDG? A")).pack()
        Button (self, text = "Back", command = lambda:self.LakeShore336MainMenu()).pack()
    def LakeShore336HeatMenu(self):
        var=0
        self.destroy()
        Frame.__init__(self)
        self.pack()
        TempLim = StringVar()
        Label (self, text = "Temperature Limit (K)").pack()
        Entry (self, textvariable = TempLim).pack()
        Button(self, text ="Send", command = lambda:self.LakeShore336("write","TLIMIT A," + TempLim.get())).pack()
        Label (self, text = "Heater Range").pack()
        Button (self, text = "High", command = lambda:self.LakeShore336("write", "RANGE 1,3")).pack()
        Button (self, text = "Medium", command = lambda:self.LakeShore336("write", "RANGE 1,2")).pack()
        Button (self, text = "Low", command = lambda:self.LakeShore336("write", "RANGE 1,1")).pack()
        Button (self, text = "OFF", command = lambda:self.LakeShore336("write", "RANGE 1,0")).pack()
        Button (self, text = "Back", command = lambda:self.LakeShore336MainMenu()).pack()
    def LakeShore336(self,option,command):
        global ans
        global kelv
        settings = open('settings.txt' , 'r')
        adress = settings.readline()
        while adress.rstrip() !='LakeShore336':
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
        global forced
        global range
        global input
        global output
        global name
        global measure
        global to
        global fr
        global tm
        range=StringVar()
        measure=StringVar()
        forced=StringVar()
        fr=StringVar()
        to=StringVar()
        name=StringVar()
        tm=StringVar()
        input = StringVar()
        output=StringVar()
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self,text='Select Automation Process').pack()
        Button(self,text='4 Wire Current vs Voltage',command=lambda:self.FourWireCurrentvsVoltaqgeMenu()).pack()
        Button(self,text='2 Wire Current vs Voltage',command=lambda:self.TwoWireCurrentvsVoltaqgeMenu()).pack()
        Button(self,text='4 Wire Voltage vs Current',command=lambda:self.FourWireVoltagevsCurrentMenu()).pack()
        Button(self,text='Voltage vs Time').pack()
        Button(self,text='Execute Process Que',command=lambda:self.UserProgramableTest1Process()).pack()
        Label(self,text='Processes in Que:').pack()
        Label(self,text=count).pack()
        Button(self,text='Back',command=lambda:self.DeviceMen()).pack()
    def FourWireVoltagevsCurrentMenu(self):
        global forced
        global count
        global range
        global input
        global output
        global name
        global measure
        global to
        global fr
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self,text='Amount forced (Volts)').pack()
        Entry(self,textvariable=forced).pack()
        Label(self,text='Range (Volts):').pack()
        Entry(self,textvariable=range).pack()
        Label(self,text='Select switch inputs From:').pack()
        Entry(self,textvariable=fr).pack()
        Label(self,text='To:').pack()
        Entry(self,textvariable=to).pack()
        Label(self,text='Name of Excel file that will be created:').pack()
        Entry(self,textvariable=name).pack()
        Button(self,text='Add this Process to Que',command=lambda:self.AddProcessToQue()).pack()
        Button(self,text='Execute Process Que',command=lambda:self.UserProgramableTest1Process()).pack()
        Label(self,text='Processes in Que:').pack()
        Label(self,text=count).pack()
        Button(self,text='Back',command = lambda:self.AutomationMenu()).pack()
        measure = '4 Wire Forced Voltage vs Current'
    def FourWireCurrentvsVoltaqgeMenu(self):
        global forced
        global count
        global range
        global input
        global output
        global name
        global measure
        global to
        global fr
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self,text='Amount forced (Amps)').pack()
        Entry(self,textvariable=forced).pack()
        Label(self,text='Range (Amps):').pack()
        Entry(self,textvariable=range).pack()
        Label(self,text='Select switch inputs From:').pack()
        Entry(self,textvariable=fr).pack()
        Label(self,text='To:').pack()
        Entry(self,textvariable=to).pack()
        Label(self,text='Name of Excel file that will be created:').pack()
        Entry(self,textvariable=name).pack()
        Button(self,text='Add this Process to Que',command=lambda:self.AddProcessToQue()).pack()
        Button(self,text='Execute Process Que',command=lambda:self.UserProgramableTest1Process()).pack()
        Label(self,text='Processes in Que:').pack()
        Label(self,text=count).pack()
        Button(self,text='Back',command = lambda:self.AutomationMenu()).pack()
        measure = '4 Wire Forced Current vs Voltage'
    def TwoWireCurrentvsVoltaqgeMenu(self):
        global forced
        global count
        global range
        global input
        global output
        global name
        global measure
        global to
        global fr
        self.destroy()
        Frame.__init__(self)
        self.pack()
        Label(self,text='Amount forced (Amps)').pack()
        Entry(self,textvariable=forced).pack()
        Label(self,text='Range (Amps):').pack()
        Entry(self,textvariable=range).pack()
        Label(self,text='Select switch inputs From:').pack()
        Entry(self,textvariable=fr).pack()
        Label(self,text='To:').pack()
        Entry(self,textvariable=to).pack()
        Label(self,text='Name of Excel file that will be created:').pack()
        Entry(self,textvariable=name).pack()
        Button(self,text='Add this Process to Que',command=lambda:self.AddProcessToQue()).pack()
        Button(self,text='Execute Process Que',command=lambda:self.UserProgramableTest1Process()).pack()
        Label(self,text='Processes in Que:').pack()
        Label(self,text=count).pack()
        Button(self,text='Back',command = lambda:self.AutomationMenu()).pack()
        measure = '2 Wire Forced Current vs Voltage'
    def AddProcessToQue(self):
        global tm
        global measure
        global fr
        global to
        global name
        global count
        global forced
        global ranges
        process = open('process_que.txt', 'a')
        process.write(str(measure)+'\n')
        process.write(str(forced.get())+'\n')
        process.write(str(range.get())+'\n')
        process.write(str(tm.get())+'\n')
        process.write(str(fr.get())+'\n')
        process.write(str(to.get())+'\n')
        process.write(str(name.get())+'\n')
        count+=1
        process.close()
        self.AutomationMenu()
    def UserProgramableTest1Process(self):
        global measure
        global tm
        global fr
        global to
        global name
        global worksheet
        global forced
        global range
        global input
        global output
        global format
        processNumber = 0
        process = open('process_que.txt', 'r') 
        while processNumber < count:
          measure = process.readline()
          forced = process.readline()
          range = process.readline()
          tm = process.readline()
          fr = process.readline()
          to = process.readline()
          name = process.readline()
          processNumber = processNumber + 1
          workbook = xlsxwriter.Workbook(str(name).rstrip()+'.xlsx')
          format=workbook.add_format()
          format.set_text_wrap()
          worksheet = workbook.add_worksheet()
          print measure.rstrip()
          self.AutoMeasure()
        workbook.close()
        self.AutomationMenu()
    def AutoMeasure(self):
        global name 
        global tm
        global measure
        global worksheet
        global fr
        global forced
        global format
        col = 0
        row = 0
        tme = 0
        self.Keithley7002('write','open all')
        if str(measure).rstrip()=='Ressistance vs Time':  # If the user checked the resistants meassurement 
            while int(fr) != int(to):
                self.Keithley7002('write','close (@1!'+(str(fr)).rstrip()+',1!10)')
                fr = int(fr)+1
                worksheet.write(row,col,'='+str(tme))
                worksheet.write(row,col+1,'='+self.Agilent34410A('ask','MEAS:RES?'))
                row+=1
                tme=tme+float(tm)
                time.sleep(float(tm))
                self.Keithley7002('write','open all')
        if str(measure.rstrip()) == '4 Wire Forced Current vs Voltage':
             worksheet.write(row,col,'Current',format)
             worksheet.write(row,col+1,'Voltage',format)
             while int(fr) < int(to)+1:
                row+=1
                fr = str(fr).rstrip()
                self.Keithley7002('write','close (@1!'+(str(fr)).rstrip()+')')
                fr = int(fr)+1
                self.YokogawaGS200('write','SENS:REM OFF')
                self.YokogawaGS200('write','SOUR:FUNC CURR')
                self.YokogawaGS200('write','SOUR:RANG '+str(range.rstrip()))
                self.YokogawaGS200('write','SOUR:LEV '+str(forced.rstrip()))
                self.YokogawaGS200('write','OUTP ON')
                time.sleep(.25)
                worksheet.write(row,col,'='+str(forced.rstrip()))
                worksheet.write(row,col+1,'='+str(self.Agilent34410A('ask','MEAS:VOLT:DC?')))
                self.YokogawaGS200('write','OUTP OFF')
                self.Keithley7002('write','open all')
        if str(measure.rstrip())=='4 Wire Forced Voltage vs Current':
            worksheet.write(row,col,'Voltage',format)
            worksheet.write(row,col+1,'Current',format)
            while int(fr) < int(to)+1:
                row+=1
                fr = str(fr).rstrip()
                self.Keithley7002('write','close (@1!'+(str(fr)).rstrip()+',1!'+str(int(fr)+1)+',1!'+str(input.rstrip())+',1!'+str(output.rstrip())+')')
                fr = int(fr)+2
                self.YokogawaGS200('write','SENS:REM ON')
                self.YokogawaGS200('write','SENS:TRIG IMM')
                self.YokogawaGS200('write','SOUR:FUNC VOLT')
                self.YokogawaGS200('write','SOUR:RANG '+str(range.rstrip()))
                self.YokogawaGS200('write','SOUR:LEV '+str(forced.rstrip()))
                self.YokogawaGS200('write','OUTP ON')
                time.sleep(.25)
                worksheet.write(row,col,'='+str(forced.rstrip()))
                worksheet.write(row,col+1,'='+str(self.YokogawaGS200('ask','MEAS?')))
                self.YokogawaGS200('write','OUTP OFF')
                self.Keithley7002('write','open all')
        if str(measure.rstrip())=='2 Wire Forced Current vs Voltage':
             worksheet.write(row,col,'Current',format)
             worksheet.write(row,col+1,'Voltage',format)
             while int(fr) < int(to)+1:
                row+=1
                fr = str(fr).rstrip()
                self.Keithley7002('write','close (@1!'+(str(fr)).rstrip()+',1!'+str(int(fr)+1)+',1!'+str(input.rstrip())+',1!'+str(output.rstrip())+')')
                fr = int(fr)+2
                self.YokogawaGS200('write','SENS:REM OFF')
                self.YokogawaGS200('write','SENS:TRIG IMM')
                self.YokogawaGS200('write','SOUR:FUNC CURR')
                self.YokogawaGS200('write','SOUR:RANG '+str(range.rstrip()))
                self.YokogawaGS200('write','SOUR:LEV '+str(forced.rstrip()))
                self.YokogawaGS200('write','OUTP ON')
                time.sleep(.25)
                worksheet.write(row,col,'='+str(forced.rstrip()))
                worksheet.write(row,col+1,'='+str(self.YokogawaGS200('ask','MEAS?')))
                self.YokogawaGS200('write','OUTP OFF')
                self.Keithley7002('write','open all')
root = Tk()
root.title("Measurement System GUI Alpha")
root.geometry("600x500")
app = Application(root)
root.mainloop() 
