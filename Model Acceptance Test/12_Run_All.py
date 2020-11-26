# Import modules
import os
# run scripts
#flag00 = os.system('python 00_Generate_Cases.py')
flag01 = os.system('python 01_Long_Run_Test.py')
flag02 = os.system('python 02_Deep_Fault_Test.py')
flag03 = os.system('python 03_Shallow_Fault_Test.py')
flag04 = os.system('python 04_PoC_Pref_Step_Test.py')
flag05 = os.system('python 05_PoC_Vref_Step_Test.py')
flag06 = os.system('python 06_PoC_Qref_Step_Test.py')
flag07 = os.system('python 07_Grid_Voltage_Step_Test.py')
flag08 = os.system('python 08_Voltage_Angle_Step_Test.py')
flag09 = os.system('python 09_Frequency_Ramp_Test.py')
flag10 = os.system('python 10_Low_Voltage_Ride_Through_Test.py')
flag11 = os.system('python 11_High_Voltage_Ride_Through_Test.py')