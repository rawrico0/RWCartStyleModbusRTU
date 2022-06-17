import pandas as pd
import time
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import minimalmodbus



# Tkinter interface program #
m = tk.Tk()
m.title('"CARTRIDGE STYLE" MODBUS RTU READ/WRITER')
m.resizable(False, False)
topf = tk.Frame(m).grid(row = 0, column = 0, columnspan = 10, rowspan = 1, sticky = 'nsew')
botf = tk.Frame(m).grid(row = 1, column = 0, columnspan = 10, sticky = 'nsew')
menu = tk.Menu(m)

# Definition Parameters #
# Serial Parameter #
Port = tk.StringVar()
Baud = tk.IntVar()
Timeout = tk.DoubleVar()
BSP = tk.StringVar()
Bytetk = tk.IntVar()
Bytetk.set(0)
Paritytk = tk.StringVar()
Paritytk.set('N')
Stoptk = tk.IntVar()
Stoptk.set(0)
modrtulib = tk.IntVar()
modrtulib.set(0)
# Controller Model #
Model = tk.StringVar()
Model.set('Ready')
# Address #
ADDRE = tk.IntVar()
ADDRE.set(0)
# Status #
status = tk.StringVar()
status.set('Open a Cartridge(XLSX) to proceed next step')
# 101 #
PL101 = tk.StringVar()
PL101.set('  ')
PM101 = tk.StringVar()
PM101.set('  ')
PE101 = tk.StringVar()
PE101.set('  ')
PK101 = tk.IntVar()
PK101.set(0)
# 102 #
PL102 = tk.StringVar()
PL102.set('  ')
PM102 = tk.StringVar()
PM102.set('  ')
PE102 = tk.StringVar()
PE102.set('  ')
PK102 = tk.IntVar()
PK102.set(0)
# 103 #
PL103 = tk.StringVar()
PL103.set('  ')
PM103 = tk.StringVar()
PM103.set('  ')
PE103 = tk.StringVar()
PE103.set('  ')
PK103 = tk.IntVar()
PK103.set(0)
# 103 #
PL104 = tk.StringVar()
PL104.set('  ')
PM104 = tk.StringVar()
PM104.set('  ')
PE104 = tk.StringVar()
PE104.set('  ')
PK104 = tk.IntVar()
PK104.set(0)
# 105 #
PL105 = tk.StringVar()
PL105.set('  ')
PM105 = tk.StringVar()
PM105.set('  ')
PE105 = tk.StringVar()
PE105.set('  ')
PK105 = tk.IntVar()
PK105.set(0)
# 106 #
PL106 = tk.StringVar()
PL106.set('  ')
PM106 = tk.StringVar()
PM106.set('  ')
PE106 = tk.StringVar()
PE106.set('  ')
PK106 = tk.IntVar()
PK106.set(0)
# 107 #
PL107 = tk.StringVar()
PL107.set('  ')
PM107 = tk.StringVar()
PM107.set('  ')
PE107 = tk.StringVar()
PE107.set('  ')
PK107 = tk.IntVar()
PK107.set(0)
# 108 #
PL108 = tk.StringVar()
PL108.set('  ')
PM108 = tk.StringVar()
PM108.set('  ')
PE108 = tk.StringVar()
PE108.set('  ')
PK108 = tk.IntVar()
PK108.set(0)
# 109 #
PL109 = tk.StringVar()
PL109.set('  ')
PM109 = tk.StringVar()
PM109.set('  ')
PE109 = tk.StringVar()
PE109.set('  ')
PK109 = tk.IntVar()
PK109.set(0)
# 110 #
PL110 = tk.StringVar()
PL110.set('  ')
PM110 = tk.StringVar()
PM110.set('  ')
PE110 = tk.StringVar()
PE110.set('  ')
PK110 = tk.IntVar()
PK110.set(0)
# 201 #
PL201 = tk.StringVar()
PL201.set('  ')
PM201 = tk.StringVar()
PM201.set('  ')
PE201 = tk.StringVar()
PE201.set('  ')
PK201 = tk.IntVar()
PK201.set(0)
# 202 #
PL202 = tk.StringVar()
PL202.set('  ')
PM202 = tk.StringVar()
PM202.set('  ')
PE202 = tk.StringVar()
PE202.set('  ')
PK202 = tk.IntVar()
PK202.set(0)
# 203 #
PL203 = tk.StringVar()
PL203.set('  ')
PM203 = tk.StringVar()
PM203.set('  ')
PE203 = tk.StringVar()
PE203.set('  ')
PK203 = tk.IntVar()
PK203.set(0)
# 204 #
PL204 = tk.StringVar()
PL204.set('  ')
PM204 = tk.StringVar()
PM204.set('  ')
PE204 = tk.StringVar()
PE204.set('  ')
PK204 = tk.IntVar()
PK204.set(0)
# 205 #
PL205 = tk.StringVar()
PL205.set('  ')
PM205 = tk.StringVar()
PM205.set('  ')
PE205 = tk.StringVar()
PE205.set('  ')
PK205 = tk.IntVar()
PK205.set(0)
# 206 #
PL206 = tk.StringVar()
PL206.set('  ')
PM206 = tk.StringVar()
PM206.set('  ')
PE206 = tk.StringVar()
PE206.set('  ')
PK206 = tk.IntVar()
PK206.set(0)
# 207 #
PL207 = tk.StringVar()
PL207.set('  ')
PM207 = tk.StringVar()
PM207.set('  ')
PE207 = tk.StringVar()
PE207.set('  ')
PK207 = tk.IntVar()
PK207.set(0)
# 208 #
PL208 = tk.StringVar()
PL208.set('  ')
PM208 = tk.StringVar()
PM208.set('  ')
PE208 = tk.StringVar()
PE208.set('  ')
PK208 = tk.IntVar()
PK208.set(0)
# 209 #
PL209 = tk.StringVar()
PL209.set('  ')
PM209 = tk.StringVar()
PM209.set('  ')
PE209 = tk.StringVar()
PE209.set('  ')
PK209 = tk.IntVar()
PK209.set(0)
# 210 #
PL210 = tk.StringVar()
PL210.set('  ')
PM210 = tk.StringVar()
PM210.set('  ')
PE210 = tk.StringVar()
PE210.set('  ')
PK210 = tk.IntVar()
PK210.set(0)

# MENU BAR #
# Menu bar variable #
comportmenu = tk.Menu(m, tearoff = 0)
# Menu bar 1st column #
comportmenu = tk.Menu(m, tearoff = 0)
comportmenu.add_radiobutton(label = 'COM1', variable = Port, value='COM1', font=('arial' ,10 ,'bold'))
comportmenu.add_radiobutton(label = 'COM2', variable = Port, value='COM2', font=('arial' ,10 ,'bold'))
comportmenu.add_radiobutton(label = 'COM3', variable = Port, value='COM3', font=('arial' ,10 ,'bold'))
comportmenu.add_radiobutton(label = 'COM4', variable = Port, value='COM4', font=('arial' ,10 ,'bold'))
comportmenu.add_radiobutton(label = 'COM5', variable = Port, value='COM5', font=('arial' ,10 ,'bold'))
comportmenu.add_radiobutton(label = 'COM6', variable = Port, value='COM6', font=('arial' ,10 ,'bold'))
comportmenu.add_radiobutton(label = 'COM7', variable = Port, value='COM7', font=('arial' ,10 ,'bold'))
comportmenu.add_radiobutton(label = 'COM8', variable = Port, value='COM8', font=('arial' ,10 ,'bold'))
comportmenu.add_radiobutton(label = 'COM9', variable = Port, value='COM9', font=('arial' ,10 ,'bold'))
menu.add_cascade(label="COM PORT", menu=comportmenu)

# Menu bar 2nd column #
baudmenu = tk.Menu(m, tearoff = 0)
baudmenu.add_radiobutton(label = '9600', variable = Baud, value='9600', font=('arial' ,10 ,'bold'))
baudmenu.add_radiobutton(label = '19200', variable = Baud, value='19200', font=('arial' ,10 ,'bold'))
baudmenu.add_radiobutton(label = '38400', variable = Baud, value='38400', font=('arial' ,10 ,'bold'))
baudmenu.add_radiobutton(label = '115200', variable = Baud, value='115200', font=('arial' ,10 ,'bold'))
menu.add_cascade(label="BAUD RATE", menu=baudmenu)

# Menu bar 3rd column #
timemenu = tk.Menu(m, tearoff = 0)
timemenu.add_radiobutton(label = '3.00', variable = Timeout, value='3.0', font=('arial' ,10 ,'bold'))
timemenu.add_radiobutton(label = '2.00', variable = Timeout, value='2.0', font=('arial' ,10 ,'bold'))
timemenu.add_radiobutton(label = '1.00', variable = Timeout, value='1.0', font=('arial' ,10 ,'bold'))
timemenu.add_radiobutton(label = '0.50', variable = Timeout, value='0.5', font=('arial' ,10 ,'bold'))
menu.add_cascade(label="TIME OUT", menu=timemenu)

# Menu bar 4th column #
modlibadjmenu = tk.Menu(m, tearoff = 0)
modlibadjmenu.add_radiobutton(label = '0', variable = modrtulib, value='0', font=('arial' ,10 ,'bold'))
modlibadjmenu.add_radiobutton(label = '+1', variable = modrtulib, value='+1', font=('arial' ,10 ,'bold'))
modlibadjmenu.add_radiobutton(label = '-1', variable = modrtulib, value='-1', font=('arial' ,10 ,'bold'))
menu.add_cascade(label="RTU ADJ", menu=modlibadjmenu)

# Read from Remote device)
def read_modbus(ADDRESS, PORT, BAUD, BYTE, PARITY, STOP, TIMEOUT, MODPAR, MODADJ, DEC):
    if MODPAR > 0:
        try:
            instrument = minimalmodbus.Instrument(PORT, int(ADDRESS))
            instrument.serial.baudrate = BAUD
            instrument.serial.bytesize = BYTE
            instrument.serial.parity = PARITY
            instrument.serial.stopbits = STOP
            instrument.serial.timeout = TIMEOUT
            instrument.mode = minimalmodbus.MODE_RTU
            reed = instrument.read_register((int(MODPAR)+int(MODADJ)),signed=True)
            return int(reed / DEC)
            status.set('Read done')
            time.sleep(0.3)
        except IOError:
            status.set('Read Error occured')
            return 'Read Error'
    else:
        return 'None'
        status.set('Read done')
        time.sleep(0.3)

# Write to Remote device)
def write_modbus(ADDRESS, PORT, BAUD, BYTE, PARITY, STOP, TIMEOUT, MODPAR, MODADJ, DEC, MODVALUE, TICK):
    if TICK == 1:
        if MODPAR > 0:
            try:
                instrument = minimalmodbus.Instrument(PORT, int(ADDRESS))
                instrument.serial.baudrate = BAUD
                instrument.serial.bytesize = BYTE
                instrument.serial.parity = PARITY
                instrument.serial.stopbits = STOP
                instrument.serial.timeout = TIMEOUT
                instrument.mode = minimalmodbus.MODE_RTU
                wrt = instrument.write_register(registeraddress=int(MODPAR+MODADJ), value=((int(MODVALUE))*(int(DEC))), number_of_decimals=0, functioncode=6, signed=True)
                reed = instrument.read_register((int(MODPAR)+int(MODADJ)),signed=True)
                return int(reed / DEC)
                time.sleep(0.3)
            except IOError:
                status.set('Write Error occured')
                return 'Write Error'
    else:
        return 'Skip'
            

        
## MAIN PROGRAM START ##
## OPEN FILE ##
def select_file():
    ## OPEN FILE SYNTAX, OPEN ONLY XLSX ##
    filetypes = (('Excel files', '*.xlsx'), ('Excel files', '*.xls'), ('All files', '*.*'))
    ## OPEN FILE AND DEFINE AS XLSX ##
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='..',
        filetypes=filetypes)
    ## DEFINE OPENED FILE AS VARIABLE ##
    READCSV = pd.read_excel(open(filename, "rb"))
    CSV = pd.DataFrame(READCSV)
    ## READ MODBUS BUTTON ##
    def READ485():
        ## NECESSARY VARIABLE COFIGURATION ##
        addrk = ADDRE.get()
        portk = Port.get()
        baudk = Baud.get()
        timeoutk  = Timeout.get()
        modadjk = modrtulib.get()
        ## USING read_modbus FUNCTION TO COMPLETE THE PROGRAM ##
        ## 101 ##
        status.set('Reading from remote device... 0')
        R101 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR101, modadjk, PD101)
        PE101.set(R101)
        m.update()
        ## 102 ##
        status.set('Reading from remote device... 5')
        R102 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR102, modadjk, PD102)
        PE102.set(R102)
        m.update()
        ## 103 ##
        status.set('Reading from remote device... 10')
        R103 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR103, modadjk, PD103)
        PE103.set(R103)
        m.update()
        ## 104 ##
        status.set('Reading from remote device... 15')
        R104 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR104, modadjk, PD104)
        PE104.set(R104)
        m.update()
        ## 105 ##
        status.set('Reading from remote device... 20')
        R105 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR105, modadjk, PD105)
        PE105.set(R105)
        m.update()
        ## 106 ##
        status.set('Reading from remote device... 25')
        R106 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR106, modadjk, PD106)
        PE106.set(R106)
        m.update()
        ## 107 ##
        status.set('Reading from remote device... 30')
        R107 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR107, modadjk, PD107)
        PE107.set(R107)
        m.update()
        ## 108 ##
        status.set('Reading from remote device... 35')
        R108 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR108, modadjk, PD108)
        PE108.set(R108)
        m.update()
        ## 109 ##
        status.set('Reading from remote device... 40')
        R109 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR109, modadjk, PD109)
        PE109.set(R109)
        m.update()
        ## 110 ##
        status.set('Reading from remote device... 45')
        R110 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR110, modadjk, PD110)
        PE110.set(R110)
        m.update()
        ## 201 ##
        status.set('Reading from remote device... 50')
        R201 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR201, modadjk, PD201)
        PE201.set(R201)
        m.update()
        ## 202 ##
        status.set('Reading from remote device... 55')
        R202 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR202, modadjk, PD202)
        PE202.set(R202)
        m.update()
        ## 203 ##
        status.set('Reading from remote device... 60')
        R203 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR203, modadjk, PD203)
        PE203.set(R203)
        m.update()
        ## 204 ##
        status.set('Reading from remote device... 65')
        R204 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR204, modadjk, PD204)
        PE204.set(R204)
        m.update()
        ## 205 ##
        status.set('Reading from remote device... 70')
        R205 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR205, modadjk, PD205)
        PE205.set(R205)
        m.update()
        ## 206 ##
        status.set('Reading from remote device... 75')
        R206 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR206, modadjk, PD206)
        PE206.set(R206)
        m.update()
        ## 207 ##
        status.set('Reading from remote device... 80')
        R207 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR207, modadjk, PD207)
        PE207.set(R207)
        m.update()
        ## 208 ##
        status.set('Reading from remote device... 85')
        R208 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR208, modadjk, PD208)
        PE208.set(R208)
        m.update()
        ## 209 ##
        status.set('Reading from remote device... 90')
        R209 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR209, modadjk, PD209)
        PE209.set(R209)
        m.update()
        ## 210 ##
        status.set('Reading from remote device... 95')
        R210 = read_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR210, modadjk, PD210)
        PE210.set(R210)
        m.update()
        status.set('Reading complete')
        m.update()
        
    ## WRITE MODBUS BUTTON ##
    def WRITE485():
        ## NECESSARY VARIABLE COFIGURATION ##
        addrk = ADDRE.get()
        portk = Port.get()
        baudk = Baud.get()
        timeoutk  = Timeout.get()
        modadjk = modrtulib.get()
        ## USING write_modbus FUNCTION TO COMPLETE THE PROGRAM ##
        ## 101 ##
        status.set('Writing to remote device... 0')
        E101 = PE101.get()
        K101 = PK101.get()
        W101 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR101, modadjk, PD101, E101, K101)
        PE101.set(W101)
        time.sleep(0.5)
        m.update()
        ## 102 ##
        status.set('Writing to remote device... 5')
        E102 = PE102.get()
        K102 = PK102.get()
        W102 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR102, modadjk, PD102, E102, K102)
        PE102.set(W102)
        time.sleep(0.5)
        m.update()
        ## 103 ##
        status.set('Writing to remote device... 10')
        E103 = PE103.get()
        K103 = PK103.get()
        W103 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR103, modadjk, PD103, E103, K103)
        PE103.set(W103)
        time.sleep(0.5)
        m.update()
        ## 104 ##
        status.set('Writing to remote device... 15')
        E104 = PE104.get()
        K104 = PK104.get()
        W104 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR104, modadjk, PD104, E104, K104)
        PE104.set(W104)
        time.sleep(0.5)
        m.update()
        ## 105 ##
        status.set('Writing to remote device... 20')
        E105 = PE105.get()
        K105 = PK105.get()
        W105 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR105, modadjk, PD105, E105, K105)
        PE105.set(W105)
        time.sleep(0.5)
        m.update()
        ## 106 ##
        status.set('Writing to remote device... 25')
        E106 = PE106.get()
        K106 = PK106.get()
        W106 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR106, modadjk, PD106, E106, K106)
        PE106.set(W106)
        time.sleep(0.5)
        m.update()
        ## 107 ##
        status.set('Writing to remote device... 30')
        E107 = PE107.get()
        K107 = PK107.get()
        W107 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR107, modadjk, PD107, E107, K107)
        PE107.set(W107)
        time.sleep(0.5)
        m.update()
        ## 108 ##
        status.set('Writing to remote device... 35')
        E108 = PE108.get()
        K108 = PK108.get()
        W108 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR108, modadjk, PD108, E108, K108)
        PE108.set(W108)
        time.sleep(0.5)
        m.update()
        ## 109 ##
        status.set('Writing to remote device... 40')
        E109 = PE109.get()
        K109 = PK109.get()
        W109 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR109, modadjk, PD109, E109, K109)
        PE109.set(W109)
        time.sleep(0.5)
        m.update()
        ## 110 ##
        status.set('Writing to remote device... 45')
        E110 = PE110.get()
        K110 = PK110.get()
        W110 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR110, modadjk, PD110, E110, K110)
        PE110.set(W110)
        time.sleep(0.5)
        m.update()
        ## 201 ##
        status.set('Writing to remote device... 50')
        E201 = PE201.get()
        K201 = PK201.get()
        W201 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR201, modadjk, PD201, E201, K201)
        PE201.set(W201)
        time.sleep(0.5)
        m.update()
        ## 202 ##
        status.set('Writing to remote device... 55')
        E202 = PE202.get()
        K202 = PK202.get()
        W202 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR202, modadjk, PD202, E202, K202)
        PE202.set(W202)
        time.sleep(0.5)
        m.update()
        ## 203 ##
        status.set('Writing to remote device... 60')
        E203 = PE203.get()
        K203 = PK203.get()
        W203 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR203, modadjk, PD203, E203, K203)
        PE203.set(W203)
        time.sleep(0.5)
        m.update()
        ## 204 ##
        status.set('Writing to remote device... 65')
        E204 = PE204.get()
        K204 = PK204.get()
        W204 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR204, modadjk, PD204, E204, K204)
        PE204.set(W204)
        time.sleep(0.5)
        m.update()
        ## 205 ##
        status.set('Writing to remote device... 70')
        E205 = PE205.get()
        K205 = PK205.get()
        W205 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR205, modadjk, PD205, E205, K205)
        PE205.set(W205)
        time.sleep(0.5)
        m.update()
        ## 206 ##
        status.set('Writing to remote device... 75')
        E206 = PE206.get()
        K206 = PK206.get()
        W206 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR206, modadjk, PD206, E206, K206)
        PE206.set(W206)
        time.sleep(0.5)
        m.update()
        ## 207 ##
        status.set('Writing to remote device... 80')
        E207 = PE207.get()
        K207 = PK207.get()
        W207 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR207, modadjk, PD207, E207, K207)
        PE207.set(W207)
        m.update()
        ## 208 ##
        status.set('Writing to remote device... 85')
        E208 = PE208.get()
        K208 = PK208.get()
        W208 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR208, modadjk, PD208, E208, K208)
        PE208.set(W208)
        time.sleep(0.5)
        m.update()
        ## 209 ##
        status.set('Writing to remote device... 90')
        E209 = PE209.get()
        K209 = PK209.get()
        W209 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR209, modadjk, PD209, E209, K209)
        PE209.set(W209)
        time.sleep(0.5)
        m.update()
        ## 210 ##
        status.set('Writing to remote device... 95')
        E210 = PE210.get()
        K210 = PK210.get()
        W210 = write_modbus(addrk, portk, baudk, Byte, Parity, Stop, timeoutk, PR210, modadjk, PD210, E210, K210)
        PE210.set(W210)
        time.sleep(0.5)
        m.update()
        

        status.set('Write Complete')
        m.update()
    
    # enabled read button #
    read_button = ttk.Button(topf, text='Read Device', command = READ485, state = 'normal').grid(row = 0, column = 8, columnspan = 1, sticky = 'nsew', rowspan = 2)
    # enabled write button #
    write_button = ttk.Button(topf, text='Write Device', command = WRITE485, state = 'normal').grid(row = 0, column = 9, columnspan = 1, sticky = 'nsew', rowspan = 2)
    ## IMPORT SERIAL SETTINGS FROM XLSX ##
    Model.set(str(CSV['Pro_Value'].values[0]))
    Byte = CSV['Pro_Value'].values[1]
    Bytetk.set(int(Byte))
    Parity = CSV['Pro_Value'].values[2]
    Paritytk.set(str(Parity))
    Stop = CSV['Pro_Value'].values[3]
    Stoptk.set(int(Stop))
    BSPMIX = Byte, Parity, Stop
    BSP.set(BSPMIX)
    ## IMPORT AND READ FROM XLSX ##
    PL101.set(CSV['param'].values[0])
    PL102.set(CSV['param'].values[1])
    PL103.set(CSV['param'].values[2])
    PL104.set(CSV['param'].values[3])
    PL105.set(CSV['param'].values[4])
    PL106.set(CSV['param'].values[5])
    PL107.set(CSV['param'].values[6])
    PL108.set(CSV['param'].values[7])
    PL109.set(CSV['param'].values[8])
    PL110.set(CSV['param'].values[9])
    PL201.set(CSV['param'].values[10])
    PL202.set(CSV['param'].values[11])
    PL203.set(CSV['param'].values[12])
    PL204.set(CSV['param'].values[13])
    PL205.set(CSV['param'].values[14])
    PL206.set(CSV['param'].values[15])
    PL207.set(CSV['param'].values[16])
    PL208.set(CSV['param'].values[17])
    PL209.set(CSV['param'].values[18])
    PL210.set(CSV['param'].values[19])
    PM101.set((CSV['value'].values[0]))
    PM102.set((CSV['value'].values[1]))
    PM103.set((CSV['value'].values[2]))
    PM104.set((CSV['value'].values[3]))
    PM105.set((CSV['value'].values[4]))
    PM106.set((CSV['value'].values[5]))
    PM107.set((CSV['value'].values[6]))
    PM108.set((CSV['value'].values[7]))
    PM109.set((CSV['value'].values[8]))
    PM110.set((CSV['value'].values[9]))
    PM201.set((CSV['value'].values[10]))
    PM202.set((CSV['value'].values[11]))
    PM203.set((CSV['value'].values[12]))
    PM204.set((CSV['value'].values[13]))
    PM205.set((CSV['value'].values[14]))
    PM206.set((CSV['value'].values[15]))
    PM207.set((CSV['value'].values[16]))
    PM208.set((CSV['value'].values[17]))
    PM209.set((CSV['value'].values[18]))
    PM210.set((CSV['value'].values[19]))
    ## DEFINE ENTRY BOX AS VARIABLE ##
    E101 = CSV['value'].values[0]
    PE101.set(E101)
    E102 = CSV['value'].values[1]
    PE102.set(E102)
    E103 = CSV['value'].values[2]
    PE103.set(E103)
    E104 = CSV['value'].values[3]
    PE104.set(E104)
    E105 = CSV['value'].values[4]
    PE105.set(E105)
    E106 = CSV['value'].values[5]
    PE106.set(E106)
    E107 = CSV['value'].values[6]
    PE107.set(E107)
    E108 = CSV['value'].values[7]
    PE108.set(E108)
    E109 = CSV['value'].values[8]
    PE109.set(E109)
    E110 = CSV['value'].values[9]
    PE110.set(E110)
    E201 = CSV['value'].values[10]
    PE201.set(E201)
    E202 = CSV['value'].values[11]
    PE202.set(E202)
    E203 = CSV['value'].values[12]
    PE203.set(E203)
    E204 = CSV['value'].values[13]
    PE204.set(E204)
    E205 = CSV['value'].values[14]
    PE205.set(E205)
    E206 = CSV['value'].values[15]
    PE206.set(E206)
    E207 = CSV['value'].values[16]
    PE207.set(E207)
    E208 = CSV['value'].values[17]
    PE208.set(E208)
    E209 = CSV['value'].values[18]
    PE209.set(E209)
    E210 = CSV['value'].values[19]
    PE210.set(E210)
    ## DEFINE MODBUS RTU CODE IN TO VARIABLE ##
    PR101 = int(CSV['rtu'].values[0])
    PR102 = int(CSV['rtu'].values[1])
    PR103 = int(CSV['rtu'].values[2])
    PR104 = int(CSV['rtu'].values[3])
    PR105 = int(CSV['rtu'].values[4])
    PR106 = int(CSV['rtu'].values[5])
    PR107 = int(CSV['rtu'].values[6])
    PR108 = int(CSV['rtu'].values[7])
    PR109 = int(CSV['rtu'].values[8])
    PR110 = int(CSV['rtu'].values[9])
    PR201 = int(CSV['rtu'].values[10])
    PR202 = int(CSV['rtu'].values[11])
    PR203 = int(CSV['rtu'].values[12])
    PR204 = int(CSV['rtu'].values[13])
    PR205 = int(CSV['rtu'].values[14])
    PR206 = int(CSV['rtu'].values[15])
    PR207 = int(CSV['rtu'].values[16])
    PR208 = int(CSV['rtu'].values[17])
    PR209 = int(CSV['rtu'].values[18])
    PR210 = int(CSV['rtu'].values[19])
    ## DEFINE PARAMETER VALUE'S DECIMAL POINT ##
    PD101 = int(CSV['dec'].values[0])
    PD102 = int(CSV['dec'].values[1])
    PD103 = int(CSV['dec'].values[2])
    PD104 = int(CSV['dec'].values[3])
    PD105 = int(CSV['dec'].values[4])
    PD106 = int(CSV['dec'].values[5])
    PD107 = int(CSV['dec'].values[6])
    PD108 = int(CSV['dec'].values[7])
    PD109 = int(CSV['dec'].values[8])
    PD110 = int(CSV['dec'].values[9])
    PD201 = int(CSV['dec'].values[10])
    PD202 = int(CSV['dec'].values[11])
    PD203 = int(CSV['dec'].values[12])
    PD204 = int(CSV['dec'].values[13])
    PD205 = int(CSV['dec'].values[14])
    PD206 = int(CSV['dec'].values[15])
    PD207 = int(CSV['dec'].values[16])
    PD208 = int(CSV['dec'].values[17])
    PD209 = int(CSV['dec'].values[18])
    PD210 = int(CSV['dec'].values[19])
    ## TICK BOX TO VERIFY WRITABLE ##
    if E101 == "=": L101tk = tk.Checkbutton(botf,variable = PK101,state = 'disabled').grid(row=4, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L101tk = tk.Checkbutton(botf,variable = PK101,state = 'normal').grid(row=4, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E102 == "=": L102tk = tk.Checkbutton(botf,variable = PK102,state = 'disabled').grid(row=5, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L102tk = tk.Checkbutton(botf,variable = PK102,state = 'normal').grid(row=5, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E103 == "=": L103tk = tk.Checkbutton(botf,variable = PK103,state = 'disabled').grid(row=6, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L103tk = tk.Checkbutton(botf,variable = PK103,state = 'normal').grid(row=6, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E104 == "=": L104tk = tk.Checkbutton(botf,variable = PK104,state = 'disabled').grid(row=7, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L104tk = tk.Checkbutton(botf,variable = PK104,state = 'normal').grid(row=7, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E105 == "=": L105tk = tk.Checkbutton(botf,variable = PK105,state = 'disabled').grid(row=8, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L105tk = tk.Checkbutton(botf,variable = PK105,state = 'normal').grid(row=8, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E106 == "=": L106tk = tk.Checkbutton(botf,variable = PK106,state = 'disabled').grid(row=9, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L106tk = tk.Checkbutton(botf,variable = PK106,state = 'normal').grid(row=9, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E107 == "=": L107tk = tk.Checkbutton(botf,variable = PK107,state = 'disabled').grid(row=10, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L107tk = tk.Checkbutton(botf,variable = PK107,state = 'normal').grid(row=10, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E108 == "=": L108tk = tk.Checkbutton(botf,variable = PK108,state = 'disabled').grid(row=11, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L108tk = tk.Checkbutton(botf,variable = PK108,state = 'normal').grid(row=11, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E109 == "=": L109tk = tk.Checkbutton(botf,variable = PK109,state = 'disabled').grid(row=12, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L109tk = tk.Checkbutton(botf,variable = PK109,state = 'normal').grid(row=12, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E110 == "=": L110tk = tk.Checkbutton(botf,variable = PK110,state = 'disabled').grid(row=13, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L110tk = tk.Checkbutton(botf,variable = PK110,state = 'normal').grid(row=13, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E201 == "=": L201tk = tk.Checkbutton(botf,variable = PK201,state = 'disabled').grid(row=4, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L201tk = tk.Checkbutton(botf,variable = PK201,state = 'normal').grid(row=4, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E202 == "=": L202tk = tk.Checkbutton(botf,variable = PK202,state = 'disabled').grid(row=5, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L202tk = tk.Checkbutton(botf,variable = PK202,state = 'normal').grid(row=5, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E203 == "=": L203tk = tk.Checkbutton(botf,variable = PK203,state = 'disabled').grid(row=6, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L203tk = tk.Checkbutton(botf,variable = PK203,state = 'normal').grid(row=6, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E204 == "=": L204tk = tk.Checkbutton(botf,variable = PK204,state = 'disabled').grid(row=7, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L204tk = tk.Checkbutton(botf,variable = PK204,state = 'normal').grid(row=7, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E205 == "=": L205tk = tk.Checkbutton(botf,variable = PK205,state = 'disabled').grid(row=8, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L205tk = tk.Checkbutton(botf,variable = PK205,state = 'normal').grid(row=8, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E206 == "=": L206tk = tk.Checkbutton(botf,variable = PK206,state = 'disabled').grid(row=9, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L206tk = tk.Checkbutton(botf,variable = PK206,state = 'normal').grid(row=9, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E207 == "=": L207tk = tk.Checkbutton(botf,variable = PK207,state = 'disabled').grid(row=10, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L207tk = tk.Checkbutton(botf,variable = PK207,state = 'normal').grid(row=10, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E208 == "=": L208tk = tk.Checkbutton(botf,variable = PK208,state = 'disabled').grid(row=11, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L208tk = tk.Checkbutton(botf,variable = PK208,state = 'normal').grid(row=11, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E209 == "=": L209tk = tk.Checkbutton(botf,variable = PK209,state = 'disabled').grid(row=12, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L209tk = tk.Checkbutton(botf,variable = PK209,state = 'normal').grid(row=12, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    if E210 == "=": L210tk = tk.Checkbutton(botf,variable = PK210,state = 'disabled').grid(row=13, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
    else: L210tk = tk.Checkbutton(botf,variable = PK210,state = 'normal').grid(row=13, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

    Timeout.set(1.00)
    Baud.set(9600)
    ADDRE.set(1)
    modrtulib.set(0)
    status.set('Cartridge(XLSX) read, please check all parameter & value and start read & write')


    
# Top Frame #
# button and entry box row#
# open button #
open_button = ttk.Button(topf, text='Open XLSX', command=select_file).grid(row = 0, column = 0, columnspan = 1, sticky = 'nsew', rowspan = 2)
# Model Name #
Modelbox = tk.Message(topf, textvariable = Model, font = ('arial', 10, 'bold'), width = 80).grid(row=0, column=1, columnspan = 3, sticky = 'nsew', rowspan = 2)
# BSP (BYTE,STOP BIT, PARITY) LABEL #
BSPboxLabel = tk.Label(topf, text = 'Serial = ', font = ('arial', 10, 'bold'), width = 10).grid(row=0, column=4, columnspan = 1, sticky = 'E', rowspan = 2)
# BSP (BYTE,STOP BIT, PARITY) MESSAGE #
BSPbox = tk.Message(topf, textvariable = BSP, font = ('arial', 10, 'bold'), width = 30).grid(row=0, column=5, columnspan = 1, sticky = 'W', rowspan = 2)
# read button #
read_button = ttk.Button(topf, text='Read Device', state = 'disabled').grid(row = 0, column = 8, columnspan = 1, sticky = 'nsew', rowspan = 2)
# write button #
write_button = ttk.Button(topf, text='Write Device', state = 'disabled').grid(row = 0, column = 9, columnspan = 1, sticky = 'nsew', rowspan = 2)
# address box #
add_box = ttk.Entry(topf, textvariable = ADDRE, font = ('arial', 10, 'bold'), width = 10).grid(row = 0, column = 10, columnspan = 1, sticky = 'nse', rowspan = 2)
# status box #
statusbox = tk.Message(topf, textvariable = status, font = ('arial', 10, 'bold'), width = 700).grid(row = 15, column = 0, columnspan = 10, sticky = 'nsw', rowspan = 1)

# Bottom Frame #
## LABEL ##
botflabel = tk.Label(botf, text = 'Cratridge style parameter', font = ('arial', 10, 'bold'), width = 8).grid(row=2, column=0, sticky = 'EW', columnspan = 10, rowspan = 1)
## LEFT COLUMN ##
botfindexParmA = tk.Label(botf, text = 'Parameter', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
botfindexValueA = tk.Label(botf, text = 'Value', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
botfindexSetA = tk.Label(botf, text = 'Set', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
botfindexButA = tk.Label(botf, text = 'Select', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)
botfindexSpaceA = tk.Label(botf, text = ' ', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=5, sticky = 'EW', columnspan = 1, rowspan = 1)
## RIGHT COLUMN ##
botfindexParmB = tk.Label(botf, text = 'Parameter', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
botfindexValueB = tk.Label(botf, text = 'Value', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
botfindexSetB = tk.Label(botf, text = 'Set', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
botfindexButB = tk.Label(botf, text = 'Select', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)
botfindexSpaceB = tk.Label(botf, text = ' ', font = ('arial', 10, 'bold'), width = 8).grid(row=3, column=11, sticky = 'EW', columnspan = 1, rowspan = 1)

## PARAMETER / VALUE ##
## 101 LABEL ##
L101label = tk.Message(botf, textvariable = PL101, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=4, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 101 ENTRY ##
L101mes = tk.Message(botf, textvariable = PM101, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=4, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 101 VALUE ##
L101ent = tk.Entry(botf, textvariable = PE101, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=4, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 101 TICK BOX ##
L101tk = tk.Checkbutton(botf,variable = PK101).grid(row=4, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)

## 102 LABEL ##
L102label =tk.Message(botf, textvariable = PL102, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=5, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 102 ENTRY ##
L102mes = tk.Message(botf, textvariable = PM102, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=5, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 102 VALUE ##
L102ent = tk.Entry(botf, textvariable = PE102, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=5, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 102 TICK BOX ##
L102tk = tk.Checkbutton(botf,variable = PK102).grid(row=5, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)

## 103 LABEL ##
L103label = tk.Message(botf, textvariable = PL103, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=6, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 103 ENTRY ##
L103mes = tk.Message(botf, textvariable = PM103, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=6, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 103 VALUE ##
L103ent = tk.Entry(botf, textvariable = PE103, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=6, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 103 TICK BOX ##
L103tk = tk.Checkbutton(botf,variable = PK103).grid(row=6, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)

## 104 LABEL ##
L104label = tk.Message(botf, textvariable = PL104, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=7, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 104 ENTRY ##
L104mes = tk.Message(botf, textvariable = PM104, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=7, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 104 VALUE ##
L104ent = tk.Entry(botf, textvariable = PE104, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=7, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 104 TICK BOX ##
L104tk = tk.Checkbutton(botf,variable = PK104).grid(row=7, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)

## 105 LABEL ##
L105label = tk.Message(botf, textvariable = PL105, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=8, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 105 ENTRY ##
L105mes = tk.Message(botf, textvariable = PM105, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=8, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 105 VALUE ##
L105ent = tk.Entry(botf, textvariable = PE105, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=8, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 105 TICK BOX ##
L105tk = tk.Checkbutton(botf,variable = PK105).grid(row=8, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)


## 106 LABEL ##
L106label = tk.Message(botf, textvariable = PL106, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=9, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 106 ENTRY ##
L106mes = tk.Message(botf, textvariable = PM106, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=9, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 106 VALUE ##
L106ent = tk.Entry(botf, textvariable = PE106, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=9, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 106 TICK BOX ##
L106tk = tk.Checkbutton(botf,variable = PK106).grid(row=9, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)

## 107 LABEL ##
L107label = tk.Message(botf, textvariable = PL107, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=10, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 107 ENTRY ##
L107mes = tk.Message(botf, textvariable = PM107, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=10, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 107 VALUE ##
L107ent = tk.Entry(botf, textvariable = PE107, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=10, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 107 TICK BOX ##
L107tk = tk.Checkbutton(botf,variable = PK107).grid(row=10, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)

## 108 LABEL ##
L108label = tk.Message(botf, textvariable = PL108, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=11, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 108 ENTRY ##
L108mes = tk.Message(botf, textvariable = PM108, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=11, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 108 VALUE ##
L108ent = tk.Entry(botf, textvariable = PE108, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=11, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 108 TICK BOX ##
L108tk = tk.Checkbutton(botf,variable = PK108).grid(row=11, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)

## 109 LABEL ##
L109label = tk.Message(botf, textvariable = PL109, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=12, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 109 ENTRY ##
L109mes = tk.Message(botf, textvariable = PM109, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=12, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 109 VALUE ##
L109ent = tk.Entry(botf, textvariable = PE109, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=12, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 109 TICK BOX ##
L109tk = tk.Checkbutton(botf,variable = PK109).grid(row=12, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)


## 110 LABEL ##
L110label = tk.Message(botf, textvariable = PL110, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=13, column=0, sticky = 'EW', columnspan = 2, rowspan = 1)
## 110 ENTRY ##
L110mes = tk.Message(botf, textvariable = PM110, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=13, column=2, sticky = 'EW', columnspan = 1, rowspan = 1)
## 110 VALUE ##
L110ent = tk.Entry(botf, textvariable = PE110, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=13, column=3, sticky = 'EW', columnspan = 1, rowspan = 1)
## 110 TICK BOX ##
L110tk = tk.Checkbutton(botf,variable = PK110).grid(row=13, column=4, sticky = 'EW', columnspan = 1, rowspan = 1)

## 201 LABEL ##
L201label = tk.Message(botf, textvariable = PL201, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=4, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 201 ENTRY ##
L201mes = tk.Message(botf, textvariable = PM201, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=4, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 201 VALUE ##
L201ent = tk.Entry(botf, textvariable = PE201, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=4, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 201 TICK BOX ##
L201tk = tk.Checkbutton(botf,variable = PK201).grid(row=4, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

## 202 LABEL ##
L202label = tk.Message(botf, textvariable = PL202, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=5, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 202 ENTRY ##
L202mes = tk.Message(botf, textvariable = PM202, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=5, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 202 VALUE ##
L202ent = tk.Entry(botf, textvariable = PE202, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=5, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 202 TICK BOX ##
L202tk = tk.Checkbutton(botf,variable = PK202).grid(row=5, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

## 203 LABEL ##
L203label = tk.Message(botf, textvariable = PL203, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=6, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 203 ENTRY ##
L203mes = tk.Message(botf, textvariable = PM203, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=6, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 203 VALUE ##
L203ent = tk.Entry(botf, textvariable = PE203, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=6, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 203 TICK BOX ##
L203tk = tk.Checkbutton(botf,variable = PK203).grid(row=6, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

## 204 LABEL ##
L204label = tk.Message(botf, textvariable = PL204, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=7, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 204 ENTRY ##
L204mes = tk.Message(botf, textvariable = PM204, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=7, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 204 VALUE ##
L204ent = tk.Entry(botf, textvariable = PE204, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=7, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 204 TICK BOX ##
L204tk = tk.Checkbutton(botf,variable = PK204).grid(row=7, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

## 205 LABEL ##
L205label = tk.Message(botf, textvariable = PL205, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=8, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 205 ENTRY ##
L205mes = tk.Message(botf, textvariable = PM205, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=8, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 205 VALUE ##
L205ent = tk.Entry(botf, textvariable = PE205, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=8, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 205 TICK BOX ##
L205tk = tk.Checkbutton(botf,variable = PK205).grid(row=8, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

## 206 LABEL ##
L206label = tk.Message(botf, textvariable = PL206, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=9, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 206 ENTRY ##
L206mes = tk.Message(botf, textvariable = PM206, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=9, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 206 VALUE ##
L206ent = tk.Entry(botf, textvariable = PE206, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=9, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 206 TICK BOX ##
L206tk = tk.Checkbutton(botf,variable = PK206).grid(row=9, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

## 207 LABEL ##
L207label = tk.Message(botf, textvariable = PL207, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=10, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 207 ENTRY ##
L207mes = tk.Message(botf, textvariable = PM207, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=10, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 207 VALUE ##
L207ent = tk.Entry(botf, textvariable = PE207, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=10, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 207 TICK BOX ##
L207tk = tk.Checkbutton(botf,variable = PK207).grid(row=10, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

## 208 LABEL ##
L208label = tk.Message(botf, textvariable = PL208, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=11, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 208 ENTRY ##
L208mes = tk.Message(botf, textvariable = PM208, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=11, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 208 VALUE ##
L208ent = tk.Entry(botf, textvariable = PE208, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=11, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 208 TICK BOX ##
L208tk = tk.Checkbutton(botf,variable = PK208).grid(row=11, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

## 209 LABEL ##
L209label = tk.Message(botf, textvariable = PL209, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=12, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 209 ENTRY ##
L209mes = tk.Message(botf, textvariable = PM209, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=12, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 209 VALUE ##
L209ent = tk.Entry(botf, textvariable = PE209, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=12, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 209 TICK BOX ##
L209tk = tk.Checkbutton(botf,variable = PK209).grid(row=12, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

## 210 LABEL ##
L210label = tk.Message(botf, textvariable = PL210, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 80).grid(row=13, column=6, sticky = 'EW', columnspan = 2, rowspan = 1)
## 210 ENTRY ##
L210mes = tk.Message(botf, textvariable = PM210, font = ('arial', 10, 'bold'), bd = 2, bg = 'white', relief= 'sunken', width = 50).grid(row=13, column=8, sticky = 'EW', columnspan = 1, rowspan = 1)
## 210 VALUE ##
L210ent = tk.Entry(botf, textvariable = PE210, font = ('arial', 10, 'bold'), bd = 2, width = 10).grid(row=13, column=9, sticky = 'EW', columnspan = 1, rowspan = 1)
## 210 TICK BOX ##
L210tk = tk.Checkbutton(botf,variable = PK210).grid(row=13, column=10, sticky = 'EW', columnspan = 1, rowspan = 1)

# Tkinter interface program (End) #
m.config(menu=menu)
m.mainloop()