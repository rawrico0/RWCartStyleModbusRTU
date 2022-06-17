# RWCartStyleModbusRTU
To Read / Write parameters through Modbus RTU
Introductions
===================
Hello internet, 
I'm new to coding, so people who review my code might found funny and fuzzy about it, 
however, it work though...

Program work as...
===================
Right, long story short,
Python script i wrote is to help read or write modbus RTU parameter from device. 
Devices such as Danfoss or Dixell, with some knowledge of Modbus RTU basics...

Reason i'm made it
===================
I've been searching for programs that aim for specific brands or modbus devices,
however, none of them are works as i wish to, or rather it could cost me many.
So i made one that can program a specific controller type.
And so, i'm think of it, i will made one that can use as "Cartridge Style", 
which all the modbus RTU parameter are corrently write on a XLSX file, 
and use this python sctipt to do the read and write job.

Requirement to run this Python Script
===================
-Your own Modbus RTU parameter code. All devices doesn't use the same parameters, you need to find it by your own. 
I've uploaded 2 Excel files as example.

-Any USB to RS485 dongle should do, in case you can't read anything, please connect a 120ohm resistor at + and - terminal, and try again.
USB Dongle i've tried... A Black USB dongle with 2 little + - terminal on it; Aten USB to RS232 dongle with a RS232-to-RS485 converter.

This Python 3.10 script required to import the followings, besure to pip them:
-tkinter
-openpyxl
-pandas
-minimalmodbus
-pyserial

Some "how to use" things will guide through in screenshot i've uploaded.
How these script helps many modbus related industries.
