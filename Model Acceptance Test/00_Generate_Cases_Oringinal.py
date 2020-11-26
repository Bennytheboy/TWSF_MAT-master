# Import modules
import glob, os, sys, math, csv, time, logging, traceback, exceptions
import shutil
# from win32com import client
# from subprocess import Popen

WorkingFolder = os.getcwd()

PSSE_LOCATION = r"C:\Program Files (x86)\PTI\PSSE34\PSSBIN"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['path'] + ';' + PSSE_LOCATION

PSSE_LOCATION = r"C:\Program Files (x86)\PTI\PSSE34\PSSPY27"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['path'] + ';' + PSSE_LOCATION

import psse34
import psspy
import redirect
import dyntools
import numpy

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
Thevenin_R = 0.02300
Thevenin_X = 0.09700
psspy.branch_chng_3(107, 969, r"""1""", [_i, _i, _i, _i, _i, _i],
                    [Thevenin_R, Thevenin_X, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f],
                    [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f], "")
psspy.fnsl([1, 0, 0, 1, 1, 0, 0, 0])

MBASE = 83.6      # MBASE = 83.6 MVA
SBASE = 100.0     # SBASE = 100 MVA
bus_inf = 969
bus_gen = 100
bus_poc = 106
bus_flt = 107
PMAX = 65.6
PMIN = 3.25
PPOC = 65
QPOC = 0.3*PPOC
POC_VCtrl_Tgt = 1.0
MisMatch_Tol = 0.001

SCR = [9.05, 9.05, 9.05,  3.0,  3.0, 3.0, 9.05, 9.05, 9.05,  3.0,
        3.0,  3.0, 9.05, 9.05, 9.05, 3.0,  3.0,  3.0, 9.05, 9.05,
       9.05,  3.0,  3.0,  3.0]

XR_ratio = [2.24, 2.24, 2.24, 2.24, 2.24, 2.24, 2.24, 2.24, 2.24, 2.24,
            2.24, 2.24, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0,
            10.0, 10.0, 10.0, 10.0]

Pgen = [PMAX, PMAX, PMAX, PMAX, PMAX, PMAX, PMIN, PMIN, PMIN, PMIN,
        PMIN, PMIN, PMAX, PMAX, PMAX, PMAX, PMAX, PMAX, PMIN, PMIN,
        PMIN, PMIN, PMIN, PMIN]

Qpoc = [ 0*QPOC,  1*QPOC, -1*QPOC,  0*QPOC,  1*QPOC, -1*QPOC, 0*QPOC,  1*QPOC, -1*QPOC, 0*QPOC,
         1*QPOC, -1*QPOC,  0*QPOC,  1*QPOC, -1*QPOC,  0*QPOC, 1*QPOC, -1*QPOC,  0*QPOC, 1*QPOC,
        -1*QPOC,  0*QPOC,  1*QPOC, -1*QPOC]

Rs = 0  # UUT source resistance
Xs = 999  # UUT source reactance

for case in range(1, 25):
    P = Pgen[case-1]
    scr = SCR[case-1]
    xr = XR_ratio[case-1]
    if Pgen[case-1] == PMAX:
        case_id = "_scr" + str(round(scr, 2)) + "_xr" + str(xr) + "_P65.0"
    else:
        case_id = "_scr" + str(round(scr, 2)) + "_xr" + str(xr) + "_P3.25"
    # Direct conversion of SCR to impedance
    # The impedance is on MBASE
    Rsys = (math.sqrt(((1.0 / scr) ** 2) / (xr ** 2 + 1.0)))  # make sure that the division is forced to be a floating point number
    Xsys = Rsys * xr
    # Convert impedances from MBASE to SBASE for entry in PSS/E
    Rsys = Rsys * (SBASE / MBASE)
    Xsys = Xsys * (SBASE / MBASE)

    psspy.fnsl([1, 0, 0, 1, 1, 0, 0, 0])
    psspy.machine_data_2(bus_gen, '1', realar1 = Pgen[case-1])  # Update the UUT active power output
    psspy.machine_data_2(bus_inf, '1', realar8 = 0.0, realar9 = 0.0001)  # Update the infinite bus equivalent impedance
    psspy.seq_machine_data(bus_inf, '1', [0.0, 0.0001, 0.0, 0.0001, _f, _f])  # Update the infinite bus equivalent impedance in sequence info
    # psspy.branch_data(bus_poc, bus_flt, '1', realar1 = Rsys * 0.1, realar2 = Xsys * 0.1)
    # psspy.branch_data(bus_flt, bus_inf, '1', realar1 = Rsys * 0.9, realar2 = Xsys * 0.9)
    psspy.branch_data(bus_flt, bus_inf, '1', realar1 = Rsys, realar2 = Xsys)
    psspy.fdns([1, 0, 1, 1, 0, 0, 99, 0])

    if case%3 == 1:
        # Qzero
        ierr, rval = psspy.brnmsc(105, bus_poc, "1", 'Q')
        ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
        Q_target = 15
        print "Line Flow_at_PoC is %s" % rval
        while abs(rval - Qpoc[case-1]) > 0.01:
            Q_target = Q_target - 0.01
            psspy.machine_data_2(bus_gen, '1', realar2=Q_target, realar3=Q_target, realar4=Q_target)
            psspy.fdns([1, 0, 1, 1, 0, 0, 99, 0])
            ierr, Vinf = psspy.busdat(bus_inf, 'PU')
            ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
            dV_POC = POC_VCtrl_Tgt - Vpoc
            while abs(dV_POC) > MisMatch_Tol:
                Vsch = Vinf + dV_POC / 2
                psspy.plant_data(bus_inf,
                                 realar1=Vsch)  # set the infinite bus scheduled voltage to the estimated voltage for this condition
                psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])
                ierr, Vinf = psspy.busdat(bus_inf, 'PU')
                ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
                dV_POC = POC_VCtrl_Tgt - Vpoc
            ierr, rval = psspy.brnmsc(105, bus_poc, "1", 'Q')
        psspy.machine_data_2(bus_gen, '1', realar2=Q_target, realar3=Q_target, realar4=Q_target)
        print "Line Flow_at_PoC is %s" % Q_target
        psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])
        # print "Vsched_for_Qzero is %s" % Vsch
        psspy.save(CasePath + 'Case' + ("%03d" % case) + str(case_id) + '_' + 'Qzero' + '.sav')

    if case%3 == 2:
        # Qlag UUT exporting reactive power from the grid
        ierr, rval = psspy.brnmsc(105, bus_poc, "1", 'Q')
        ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
        Q_target = 30
        print "Line Flow_at_PoC is %s" % rval
        while abs(rval - Qpoc[case-1]) > 0.01:
            Q_target = Q_target - 0.01
            psspy.machine_data_2(bus_gen, '1', realar2=Q_target, realar3=Q_target, realar4=Q_target)
            psspy.fdns([1, 0, 1, 1, 0, 0, 99, 0])
            ierr, Vinf = psspy.busdat(bus_inf, 'PU')
            ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
            dV_POC = POC_VCtrl_Tgt - Vpoc
            while abs(dV_POC) > MisMatch_Tol:
                Vsch = Vinf + dV_POC / 2
                psspy.plant_data(bus_inf,
                                 realar1=Vsch)  # set the infinite bus scheduled voltage to the estimated voltage for this condition
                psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])
                ierr, Vinf = psspy.busdat(bus_inf, 'PU')
                ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
                dV_POC = POC_VCtrl_Tgt - Vpoc
            ierr, rval = psspy.brnmsc(105, bus_poc, "1", 'Q')
        psspy.machine_data_2(bus_gen, '1', realar2=Q_target, realar3=Q_target, realar4=Q_target)
        print "Line Flow_at_PoC is %s" % Q_target
        psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])
        # print "Vsched_for_Qlag is %s" % Vsch
        psspy.save(CasePath + 'Case' + ("%03d" % case) + str(case_id) + '_' + 'Qlag' + '.sav')

    if case%3 == 0:
        # Qlead UUT importing reactive power from the grid
        ierr, rval = psspy.brnmsc(105, bus_poc, "1", 'Q')
        ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
        Q_target = -5
        print "Line Flow_at_PoC is %s" % rval
        while abs(rval - Qpoc[case-1]) > 0.01:
            Q_target = Q_target - 0.01
            psspy.machine_data_2(bus_gen, '1', realar2=Q_target, realar3=Q_target, realar4=Q_target)
            psspy.fdns([1, 0, 1, 1, 0, 0, 99, 0])
            ierr, Vinf = psspy.busdat(bus_inf, 'PU')
            ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
            dV_POC = POC_VCtrl_Tgt - Vpoc
            while abs(dV_POC) > MisMatch_Tol:
                Vsch = Vinf + dV_POC / 2
                psspy.plant_data(bus_inf,
                                 realar1=Vsch)  # set the infinite bus scheduled voltage to the estimated voltage for this condition
                psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])
                ierr, Vinf = psspy.busdat(bus_inf, 'PU')
                ierr, Vpoc = psspy.busdat(bus_poc, 'PU')
                dV_POC = POC_VCtrl_Tgt - Vpoc
            ierr, rval = psspy.brnmsc(105, bus_poc, "1", 'Q')
        psspy.machine_data_2(bus_gen, '1', realar2=Q_target, realar3=Q_target, realar4=Q_target)
        print "Line Flow_at_PoC is %s" % Q_target
        psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])
        # print "Vsched_for_Qlead is %s" % Vsch
        psspy.save(CasePath + 'Case' + ("%03d" % case) + str(case_id) + '_' + 'Qlead' + '.sav')