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

ProgramPath = WorkingFolder + '\\'
CasePath = WorkingFolder + "\\" + 'Cases' + "\\"
FigurePath = WorkingFolder + "\\" + 'Results' + "\\"
f_list = os.listdir(CasePath)

for test in range(1, 9):

    case = [3,  6,  9, 12, 15, 18, 21, 24]
    n_iteration = 100  # Number of max. iterations
    acceleration = 0.20  # acceleration sfactor used in the network solution
    tolerance = 0.0001  # Convergence Tolerance
    integration_step = 0.001  # Simulation time step[s]
    bus_flt = 107
    bus_inf = 969
    bus_IDTRF = 108

    # OPEN PSSE
    _i = psspy.getdefaultint()
    _f = psspy.getdefaultreal()
    _s = psspy.getdefaultchar()
    redirect.psse2py()
    psspy.psseinit(50000)

    FileName = f_list[case[test - 1]-1]
    psspy.case(CasePath + FileName)
    psspy.resq(CasePath + "Tamworth_SMIB.seq")
    psspy.addmodellibrary(CasePath + 'SMASC_E161_SMAPPC_E130_344_IVF150.dll')
    psspy.dyre_new([1, 1, 1, 1], CasePath + "Tamworth_SMIB_E161_E130_LVRT.dyr", '', '', '')
    #psspy.branch_data(bus_flt, bus_inf, '1', realar1=0, realar2=0.0001)
    #ierr, line_rx = psspy.brndt2(bus_flt, bus_inf, '1', 'RX')
    OutputFilePath = FigurePath + '10. Low Voltage Ride Through' + "\\" + 'Test0' + str(
        test) + '_' + FileName + '_Low Voltage Ride Through.out'
    # ! Setup Dynamic Simulation parameters
    psspy.dynamics_solution_param_2([n_iteration, _i, _i, _i, _i, _i, _i, _i],
                                    [acceleration, tolerance, integration_step, _f, _f, _f, _f, _f])
    psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])

    # convert load , do not change
    psspy.cong(0)
    psspy.conl(0, 1, 1, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.conl(0, 1, 2, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.conl(0, 1, 3, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.ordr(0) # ! Order the matrix: ORDR
    psspy.fact()  # ! Factorize the matrix: FACT
    psspy.tysl(0) # ! TYSL

    psspy.bus_frequency_channel([1, 969], r"""System frequency""")
    psspy.voltage_channel([2, -1, -1, 969], r"""IB_Voltage""")
    psspy.voltage_channel([3, -1, -1, 100], r"""UUT_Voltage""")
    psspy.voltage_channel([4, -1, -1, 106], r"""POC_Voltage""")
    psspy.machine_array_channel([5, 2, 100], r"""1""", r"""UUT_Pelec""")
    psspy.machine_array_channel([6, 3, 100], r"""1""", r"""UUT_Qelec""")
    psspy.branch_p_and_q_channel([7, -1, -1, 105, 106], r"""1""", [r"""POC_Flow""", ""])
    psspy.machine_array_channel([9, 9, 100], r"""1""", r"""UUT_IDcmd""")
    psspy.machine_array_channel([10, 12, 100], r"""1""", r"""UUT_IQcmd""")
    psspy.machine_array_channel([11, 8, 100], r"""1""", r"""PPC_Pcmd""")
    psspy.machine_array_channel([12, 5, 100], r"""1""", r"""PPC_Qcmd""")
    [ierr, var_sc_setp] = psspy.mdlind(100, '1', 'GEN', 'VAR')
    psspy.var_channel([13, var_sc_setp + 115], "LVRT Flag")
    psspy.var_channel([14, var_sc_setp + 116], "HVRT Flag")

    psspy.strt_2([0, 0], OutputFilePath)
    psspy.run(0, 80.0, 1000, 5, 5)