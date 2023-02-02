import xlsxwriter
import numpy as np
from matplotlib import pyplot as plt


# DATA STORED AS V T R F P in dictionary!!!

class PulseGenerator:

    def __init__(self):
        self.pulseDict = {}
        self.numPulses = 0
        self.baseParams = []
        self.baseVolt = 0
        self.baseTime = 0
        self.baseRise = 0
        self.baseFall = 0
        self.basePause = 0

        self.times = [0]
        self.voltages = [0]
        self.totaltime = 0
        self.intervals = 0
        self.interval_time = 0
        self.average_time = 0

    def setPulses(self, pulses):
        self.numPulses = pulses

    def defineBasePulse(self, volt, time, rise, fall, pause):
        self.baseParams = [volt, time, pause, rise, fall]
        for i in range(self.numPulses):
            self.pulseDict[str(i + 1)] = self.baseParams

    def checkPulse(self, i):
        var = str(i)
        if var in self.pulseDict.keys():
            print(self.pulseDict[var])
        else:
            print('Incorrect Value')

    def editPulse(self, i, v=0, t=0, p=0):
        if v:
            self.pulseDict[str(i)][0] = v
        if t:
            self.pulseDict[str(i)][1] = t
        if p:
            self.pulseDict[str(i)][2] = p

    def getDict(self):
        print(self.pulseDict)


class NeuroPulse:

    def __init__(self):
        self.numPulses = 4
        self.pulseRise = 1e-5
        self.pulseFall = 1e-5
        self.init_pause = 0

        self.baseParamsWrite = [1, 0.2, 1e-5, 1e-5, 1]  # [Volt, Time, Rise, Fall, Pause]
        self.baseParamsRead = [0.1, 0.2, 1e-5, 1e-5, 1]

        self.pulseDict = {}

        self.times = [0]
        self.volts = [0]

    def setBaseParams(self, R, F, P, T, Vw, Vr):
        write_list = [float(Vw), float(T), float(R), float(F), float(P)]
        read_list = [float(Vr), float(T), float(R), float(F), float(P)]
        self.baseParamsWrite = write_list
        self.baseParamsRead = read_list

    def viewBaseParams(self):
        print('Volt - Time - Rise - Fall - Pause')
        print('Write Pulse = ', self.baseParamsWrite)
        print('Read Pulse = ', self.baseParamsRead)

    def buildDict(self):

        baseParamsErase = [i for i in self.baseParamsWrite]
        baseParamsErase[0] *= -1
        self.pulseDict['1'] = self.baseParamsWrite
        self.pulseDict['2'] = self.baseParamsRead
        self.pulseDict['3'] = baseParamsErase
        self.pulseDict['4'] = self.baseParamsRead

    def editPulse(self, pulsenum, index, value):
        if str(pulsenum) in self.pulseDict.keys():
            self.pulseDict[str(pulsenum)][index] = value
        else:
            return 0

    def getDict(self):
        return self.pulseDict

    def setInitPause(self, init_time):
        self.init_pause = init_time
        self.times.append(init_time)
        self.volts.append(0)

    def buildArray(self, repeats=1):
        time = self.times[-1]
        for i in range(repeats):
            for value in self.pulseDict.values():
                time += value[2]
                self.times.append(time)
                self.volts.append(value[0])
                time += value[1]
                self.times.append(time)
                self.volts.append(value[0])
                time += value[3]
                self.times.append(time)
                self.volts.append(0)
                time += value[4]
                self.times.append(time)
                self.volts.append(0)
        return self.times, self.volts

    def clearArray(self):
        self.times = [0, self.init_pause]
        self.volts = [0, 0]

    def viewPulse(self):
        plt.plot(self.times, self.volts)
        plt.show()


class VStep:

    def __init__(self):
        self.numPulses = 5
        self.vsteps = [1, 2, 3, 4, 5]
        self.init_pause = 0

        self.baseParamsWrite = [1, 0.2, 1e-5, 1e-5, 1]  # [Volt, Time, Rise, Fall, Pause]
        self.baseParamsRead = [0.1, 0.2, 1e-5, 1e-5, 1]

        self.pulseDict = {}

        self.times = [0]
        self.volts = [0]

    def setBaseParams(self, R, F, P, T):
        write_list = [float(T), float(R), float(F), float(P)]
        read_list = [float(T), float(R), float(F), float(P)]
        self.baseParamsWrite[1:] = write_list
        self.baseParamsRead[1:] = read_list

    def viewBaseParams(self):
        print('Volt - Time - Rise - Fall - Pause')
        print('Write Pulse = ', self.baseParamsWrite)
        print('Read Pulse = ', self.baseParamsRead)

    def setSteps(self, start, step, num):
        self.numPulses = num
        a = []
        for i in range(num):
            a.append(start + (i * step))
        self.vsteps = a

    def buildDict(self):

        for i in range(self.numPulses):
            self.pulseDict[str(i + 1)] = [self.vsteps[i]] + self.baseParamsWrite[1:]

    def editPulse(self, pulsenum, index, value):
        if str(pulsenum) in self.pulseDict.keys():
            self.pulseDict[str(pulsenum)][index] = value
        else:
            return 0

    def addPulse(self, arr):
        self.numPulses += 1
        self.pulseDict[str(self.numPulses)] = arr

    def getDict(self):
        return self.pulseDict

    def setInitPause(self, init_time):
        self.init_pause = init_time
        self.times.append(init_time)
        self.volts.append(0)

    def buildArray(self, repeats=1):
        time = self.times[-1]
        for i in range(repeats):
            for value in self.pulseDict.values():
                time += value[2]
                self.times.append(time)
                self.volts.append(value[0])
                time += value[1]
                self.times.append(time)
                self.volts.append(value[0])
                time += value[3]
                self.times.append(time)
                self.volts.append(0)
                time += value[4]
                self.times.append(time)
                self.volts.append(0)
        return self.times, self.volts

    def clearArray(self):
        self.times = [0, self.init_pause]
        self.volts = [0, 0]

    def viewPulse(self):
        plt.plot(self.times, self.volts)
        plt.show()


class TStep:

    def __init__(self):
        self.numPulses = 5
        self.tsteps = [1, 2, 3, 4, 5]
        self.init_pause = 0

        self.baseParamsWrite = [1, 1e-5, 1e-5, 1]  # [Volt, Rise, Fall, Pause]
        self.baseParamsRead = [0.1, 1e-5, 1e-5, 1]

        self.pulseDict = {}

        self.times = [0]
        self.volts = [0]

    def setBaseParams(self, R, F, P, V):
        write_list = [float(V), float(R), float(F), float(P)]
        read_list = [float(V), float(R), float(F), float(P)]
        self.baseParamsWrite = write_list
        self.baseParamsRead = read_list

    def viewBaseParams(self):
        print('Volt - Time - Rise - Fall - Pause')
        print('Write Pulse = ', self.baseParamsWrite)
        print('Read Pulse = ', self.baseParamsRead)

    def setSteps(self, start, step, num):
        self.numPulses = num
        a = []
        for i in range(num):
            a.append(start + (i * step))
        self.tsteps = a

    def buildDict(self):

        for i in range(self.numPulses):
            self.pulseDict[str(i + 1)] = [self.baseParamsWrite[0]] + [self.tsteps[i]] + self.baseParamsWrite[1:]

    def editPulse(self, pulsenum, index, value):
        if str(pulsenum) in self.pulseDict.keys():
            self.pulseDict[str(pulsenum)][index] = value
        else:
            return 0

    def addPulse(self, arr):
        self.numPulses += 1
        self.pulseDict[str(self.numPulses)] = arr

    def getDict(self):
        return self.pulseDict

    def setInitPause(self, init_time):
        self.init_pause = init_time
        self.times.append(init_time)
        self.volts.append(0)

    def buildArray(self, repeats=1):
        time = self.times[-1]
        for i in range(repeats):
            for value in self.pulseDict.values():
                time += value[1]
                self.times.append(time)
                self.volts.append(value[0])
                time += value[1]
                self.times.append(time)
                self.volts.append(value[0])
                time += value[3]
                self.times.append(time)
                self.volts.append(0)
                time += value[4]
                self.times.append(time)
                self.volts.append(0)
        return self.times, self.volts

    def clearArray(self):
        self.times = [0, self.init_pause]
        self.volts = [0, 0]

    def viewPulse(self):
        plt.plot(self.times, self.volts)
        plt.show()


class Custom:

    def __init__(self):
        self.numPulses = 0
        self.pulseRise = 1e-5
        self.pulseFall = 1e-5
        self.init_pause = 0

        self.baseParamsWrite = [1, 0.2, 1e-5, 1e-5, 1]  # [Volt, Time, Rise, Fall, Pause]

        self.pulseDict = {}

        self.times = [0]
        self.volts = [0]

    def setBaseParams(self, R, F, P, T, V):
        write_list = [float(V), float(T), float(R), float(F), float(P)]
        self.baseParamsWrite = write_list

    def viewBaseParams(self):
        print('Volt - Time - Rise - Fall - Pause')
        print('Write Pulse = ', self.baseParamsWrite)
        print('Read Pulse = ', self.baseParamsRead)

    def setNumPulses(self, num):
        self.numPulses = num

    def buildDictEmpty(self):
        if self.numPulses:
            for i in range(self.numPulses):
                self.pulseDict[str(i + 1)] = []
        else:
            self.pulseDict = {}

    def editPulseInd(self, pulse, arr):
        if pulse in self.pulseDict.keys():
            self.pulseDict[str(pulse)] = arr
        else:
            return 0

    def buildDict(self):

        for i in range(self.numPulses):
            self.pulseDict[str(i + 1)] = self.baseParamsWrite

    def editPulse(self, pulsenum, index, value):
        if str(pulsenum) in self.pulseDict.keys():
            self.pulseDict[str(pulsenum)][index] = value
        else:
            return 0

    def addPulse(self, arr):
        self.numPulses += 1
        self.pulseDict[str(self.numPulses)] = arr

    def getDict(self):
        return self.pulseDict

    def setInitPause(self, init_time):
        self.init_pause = init_time
        self.times.append(init_time)
        self.volts.append(0)

    def buildArray(self, repeats=1):
        time = self.times[-1]
        for i in range(repeats):
            for value in self.pulseDict.values():
                time += value[2]
                self.times.append(time)
                self.volts.append(value[0])
                time += value[1]
                self.times.append(time)
                self.volts.append(value[0])
                time += value[3]
                self.times.append(time)
                self.volts.append(0)
                time += value[4]
                self.times.append(time)
                self.volts.append(0)
        return self.times, self.volts

    def clearArray(self):
        self.times = [0, self.init_pause]
        self.volts = [0, 0]

    def viewPulse(self):
        plt.plot(self.times, self.volts)
        plt.show()


def start():
    while True:
        print('Choose pulse structure:')
        print('1 - Neuromorphic')
        print('2 - V Steps')
        print('3 - T Steps')
        print('4 - Custom')
        choice = input('Enter Choice: ')
        choices = [1, 2, 3, 4]

        if choice.isnumeric():
            if int(choice) in choices:
                return choice
                break
            else:
                print('Invalid Choice')
        else:
            print('Invalid Selection')


def initialise_class(var):
    var = int(var)
    if var == 1:
        a = NeuroPulse()
        return a
    if var == 2:
        a = VStep()
        return a
    if var == 3:
        a = TStep()
        return a
    if var == 4:
        a = Custom()
        return a


def set_init_pause(object):
    pause_time = float(input('Set initial hold time - '))
    object.setInitPause(pause_time)


# For V or T Step
def menu_num_p(choice, object):
    if choice == '2' or choice == '3':
        var1 = int(input('Set num pulses - '))
        var2 = float(input('Set start Voltage/Time - '))
        var3 = float(input('Set V/T step - '))
        object.setSteps(var2, var3, var1)


# For Custom
def build_dict_empty(choice, object):
    if choice == '4':
        object.buildDictEmpty()


def set_base_params(choice, object):
    if choice == '1':
        v1 = input('Enter Write V - ')
        v2 = input('Enter Read V - ')
        v3 = input('Enter T - ')
        v4 = input('Enter Rise/Fall - ')
        v5 = input('Enter Pause')
        object.setBaseParams(v4, v4, v5, v3, v1, v2)
    if choice == '2':
        v3 = input('Enter T - ')
        v4 = input('Enter Rise/Fall - ')
        v5 = input('Enter Pause')
        object.setBaseParams(v4, v4, v5, v3)
    if choice == '3':
        v1 = input('Enter V - ')
        v4 = input('Enter Rise/Fall - ')
        v5 = input('Enter Pause')
        object.setBaseParams(v4, v4, v5, v1)


def build_dict(object):
    object.buildDict()
    print('Pulse # : [V, T, Rise, Fall, Pause]')
    print(object.getDict())


def view_pulse(choice, object):
    if choice != '4':
        a, b = object.buildArray()
        object.viewPulse()
        object.clearArray()


def get_array():
    a = float(input('Input Voltage - '))
    b = float(input('Input Pulse Time - '))
    c = float(input('Input Rise/Fall Time - '))
    d = float(input('Input Pause Time - '))
    arr = [a, b, c, c, d]
    return arr


def get_index(choice):
    if choice in ['1', '2', '3', '4']:
        print('Voltage - 1 \n Pulse Time - 2, \n Rise - 3 \n Fall - 4 \n Pause - 5')
        var = input('Select Index to Change: ')
        return int(var) - 1


def modify_pulse(choice, object):
    if choice != '4':
        while True:
            print('1 - Edit Pulse')
            print('2 - Add Pulse to end')
            print('0 - Done.')
            a = input('Enter Option = ')

            if a == '0':
                break

            if a == '2':
                arr = get_array()
                object.addPulse(arr)

            if a == '1':
                var1 = int(input('Input pulse number to edit: '))
                var2 = get_index(choice)
                var3 = float(input('Enter new value: '))
                object.editPulse(var1, var2, var3)

    if choice == '4':
        while True:
            print('1 - Edit Single Pulse Value')
            print('2 - Edit full pulse value')
            print('2 - Add Pulse to end')
            print('0 - Done.')
            a = input('Enter Option = ')

            if a == '0':
                break

            if a == '3':
                arr = get_array()
                object.addPulse(arr)

            if a == '1':
                var1 = int(input('Input pulse number to edit: '))
                var2 = get_index()
                var3 = float(input('Enter new value: '))
                object.editPulse(var1, var2, var3)

            if a == '2':
                arr = get_array()
                var1 = int(input('Input pulse number to edit: '))
                object.editPulseInd(var1, arr)


def build_array(object):
    v1, v2 = object.buildArray()


def export_pulse(object):
    a = input('Save Pulse Y/N:')
    if a.lower() == 'y':
        name = input('Input File Name: ')
        a = str(name) + '.xlsx'
        workbook = xlsxwriter.Workbook(a)
        worksheet = workbook.add_worksheet()

        row = 0
        col = 0

        for time, volt in zip(object.times, object.volts):
            worksheet.write(row, col, time)
            worksheet.write(row, col + 1, volt)
            row += 1

        workbook.close()
    else:
        print(object.getDict())


def ask_view_pulse(choice, object):
    a = input('View Pulse Y/N')
    if a.lower() == 'y':
        view_pulse(choice, object)
    elif a.lower() != 'n':
        print('Invalid choice')
        ask_view_pulse(choice, object)


def ask_modify(choice, object):
    while True:
        print('1 - Modify Pulse')
        print('2 - View Pulse')
        print('3 - Export Pulse to File')
        print('0 - Escape')
        var = input('Select Choice: ')

        if var == '1':
            modify_pulse(choice, object)

        if var == '2':
            view_pulse(choice, object)

        if var == '3':
            export_pulse(object)
            break

        if var == '0':
            break

        else:
            print('Invalid Selection')


def run_app():
    while True:
        choice = start()
        object = initialise_class(choice)
        set_init_pause(object)
        menu_num_p(choice, object)
        build_dict(object)
        view_pulse(choice, object)
        ask_modify(choice, object)

        a = input('Run Again? Y/N')
        if a.lower() == 'n':
            break


run_app()

# Modulating pulse in Neuromorphic changes both read pusles.
# Pulse Time for Neuromorphic
# Repeats
