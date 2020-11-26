# Import modules
import glob, os, sys, math, csv, time, logging, traceback, exceptions
import shutil

WorkingFolder = os.getcwd()

import os,sys
sys_path_PSSE = r"C:\Program Files (x86)\PTI\PSSE34\PSSBIN"
os_path_PSSE = r"C:\Program Files (x86)\PTI\PSSE34\PSSPY27"
sys.path.append(sys_path_PSSE)
sys.path.append(os_path_PSSE)
os.environ['PATH'] = os.environ['PATH'] + ';' + os_path_PSSE
os.environ['PATH'] = os.environ['PATH'] + ';' + sys_path_PSSE

import psse34
import psspy
import redirect
import dyntools
import numpy as np

# Set Simulation Path.
ProgramPath = WorkingFolder + '\\'
CasePath = WorkingFolder + "\\" + 'Cases' + "\\"
FigurePath = WorkingFolder + "\\" + 'Results' + "\\"

_i = psspy.getdefaultint()
_f = psspy.getdefaultreal()
_s = psspy.getdefaultchar()
psspy.psseinit()  # initialise PSSE so psspy commands can be called

psspy.throwPsseExceptions = True
SAV_File = 'ModelAcceptanceSystem.sav'
ierr = psspy.case(ProgramPath + SAV_File)

# Add a branch between the DUMMY node and the Infinite bus
Thevenin_R = 0.0000
Thevenin_X = 0.0001
psspy.branch_chng_3(107, 969, r'1', [_i, _i, _i, _i, _i, _i],
                    [Thevenin_R, Thevenin_X, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f],
                    [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f], "")
psspy.fnsl([1, 0, 0, 1, 1, 0, 0, 0])

MBASE = 83.6                # MBASE = 83.6 MVA
SBASE = 100.0               # SBASE = 100 MVA
bus_inf = 969               # infinite bus
bus_gen = 100               # inverter bus
bus_poc = 106               # point of connection
bus_dummy = 107             # dummy transformer
poc_p_max = 65              # active power at poc
poc_p_min = poc_p_max*0.05  # active power at poc
poc_q = 0.3*poc_p_max        # reactive power at poc
POC_VCtrl_Tgt = 1.0         # controlled voltage at poc
Vinf_MisMatch_Tol = 0.0001
PQMisMatch_Tol = Vinf_MisMatch_Tol*1
QAdjust_Step = Vinf_MisMatch_Tol*1
P_factor = 2
V_factor = 2
Q_factor = 2

QgenINImax = [15, 30, 0]
QgenINImin = [0, 0, -30]

SCR_Set = [9.05, 3.0]
XR_ratio_Set = [2.24, 10.0]
Ppoc_Set = [poc_p_max, poc_p_min]
Qpoc_Set = [0, poc_q, -poc_q]

scenario_Set = zip(*[[xrratio, ppoc, scr, qpoc]
              for xrratio in XR_ratio_Set
              for ppoc in Ppoc_Set
              for scr in SCR_Set
              for qpoc in Qpoc_Set])
SCR = list(scenario_Set[2])
XR_ratio = list(scenario_Set[0])
Ppoc = list(scenario_Set[1])
Qpoc = list(scenario_Set[3])

Rs = 0  # UUT source resistance
Xs = 999  # UUT source reactance

VpocALL = np.ones(24)*999
QpocALL = np.ones(24)*999
PpocALL = np.ones(24)*999

for caseNum in range(17,24):
    ppoc = Ppoc[caseNum]
    scr = SCR[caseNum]
    xrratio = XR_ratio[caseNum]
    qpoc = Qpoc[caseNum]
    caseIndex = caseNum%3
    if ppoc == poc_p_max:
        case_id = "_scr" + str(round(scr, 2)) + "_xr" + str(xrratio) + "_P65.0"
    else:
        case_id = "_scr" + str(round(scr, 2)) + "_xr" + str(xrratio) + "_P3.25"
    # Direct conversion of SCR to impedance # The impedance is on MBASE
    Rsys = math.sqrt(((1.0 / scr) ** 2) / (xrratio ** 2 + 1.0))  # make sure that the division is forced to be a floating point number
    Xsys = Rsys * xrratio
    # Convert impedances from MBASE to SBASE for entry in PSS/E
    Rsys = Rsys * (SBASE / poc_p_max)
    Xsys = Xsys * (SBASE / poc_p_max)

    psspy.fnsl([1, 0, 0, 1, 1, 0, 0, 0])

    psspy.machine_data_2(bus_gen, '1', realar1 = ppoc)  # Update the UUT active power output
    psspy.machine_data_2(bus_inf, '1', realar8 = 0.0, realar9 = 0.0001)  # Update the infinite bus equivalent impedance
    psspy.seq_machine_data(bus_inf, '1', [0.0, 0.0001, 0.0, 0.0001, _f, _f])  # Update the infinite bus equivalent impedance in sequence info
    psspy.branch_data(bus_dummy, bus_inf, '1', realar1 = Rsys, realar2 = Xsys)
    psspy.fdns([1, 0, 1, 1, 0, 0, 99, 0])
    #if caseIndex == 0:# Qexport at POC is 0
    PpocNOW = ppoc*1.01+0.1
    dPpoc = PpocNOW-ppoc
    PgenNOW = ppoc
    while abs(dPpoc) > PQMisMatch_Tol:
        PgenNOW = PgenNOW-dPpoc/P_factor
        QpocNOW = qpoc * 1.01 + 0.1
        dQpoc = QpocNOW - qpoc
        QgenNOW = QgenINImax[caseIndex]
        while abs(dQpoc) > PQMisMatch_Tol:
            QgenNOW = QgenNOW - dQpoc / Q_factor
            psspy.machine_data_2(bus_gen, '1', realar1=PgenNOW, realar2=QgenNOW, realar3=QgenNOW, realar4=QgenNOW)
            psspy.fdns([1, 0, 1, 1, 0, 0, 99, 0])
            ierr, Vinf = psspy.busdat(bus_inf, 'PU')
            ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
            dV_POC = POC_VCtrl_Tgt - Vpoc
            while abs(dV_POC) > Vinf_MisMatch_Tol:
                Vsch = Vinf + dV_POC / V_factor
                psspy.plant_data(bus_inf, realar1=Vsch)  # set the infinite bus scheduled voltage to the estimated voltage for this condition
                psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])
                ierr, Vinf = psspy.busdat(bus_inf, 'PU')
                ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
                dV_POC = POC_VCtrl_Tgt - Vpoc
            ierr, QpocNOW = psspy.brnmsc(bus_poc, bus_dummy, "1", 'Q')
            dQpoc = QpocNOW-qpoc
        ierr, PpocNOW = psspy.brnmsc(bus_poc, bus_dummy, "1", 'P')
        dPpoc = PpocNOW - ppoc
        print(dPpoc)
    psspy.machine_data_2(bus_gen, '1', realar1=PgenNOW, realar2=QgenNOW, realar3=QgenNOW, realar4=QgenNOW)
    ierr, PpocNOW = psspy.brnmsc(bus_poc, bus_dummy, "1", 'P')
    ierr, QpocNOW = psspy.brnmsc(bus_poc, bus_dummy, "1", 'Q')

    print "Line Flow_at_PoC (P) is %s" % PpocNOW
    print "Line Flow_at_PoC (Q) is %s" % QpocNOW
    psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])
    if caseIndex == 0:  # Qexport at POC is 0
        psspy.save(CasePath + 'Case' + ("%03d" % (caseNum+1)) + str(case_id) + '_' + 'Qzero' + '.sav')
    elif caseIndex == 1:
        psspy.save(CasePath + 'Case' + ("%03d" % (caseNum+1)) + str(case_id) + '_' + 'Qlag' + '.sav')
    elif caseIndex == 2:
        psspy.save(CasePath + 'Case' + ("%03d" % (caseNum+1)) + str(case_id) + '_' + 'Qlead' + '.sav')
    print(caseNum)
    VpocALL[caseNum] = Vpoc
    QpocALL[caseNum] = QpocNOW
    PpocALL[caseNum] = PpocNOW
print(VpocALL)
print(QpocALL)
print(PpocALL)