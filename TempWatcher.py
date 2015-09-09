import time

import visa

normal_conductance_voltage = '1'
super_conducting_voltage = '15'


def LakeShore336(option, command):
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


def YokogawaGS200(option, command):
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


while True:
    temp_ = float(LakeShore336("ask", "KRDG? A"))
    print temp_
    if temp_ >= 5.0:
        forced_voltage = normal_conductance_voltage
        YokogawaGS200('write', 'SOUR:RANG ' + str(forced_voltage))
        YokogawaGS200('write', 'SOUR:LEV ' + str(forced_voltage))
        prev_state = "NC"
        time.sleep(1)
    if prev_state == "NC" and temp_ < 5.0:
        forced_voltage = super_conducting_voltage
        volt_ramp = float(normal_conductance_voltage)
        prev_state = "SC"
        while volt_ramp < float(forced_voltage):
            volt_ramp += 0.5
            YokogawaGS200('write', 'SOUR:RANG ' + str(volt_ramp))
            YokogawaGS200('write', 'SOUR:LEV ' + str(volt_ramp))
            time.sleep(0.25)
