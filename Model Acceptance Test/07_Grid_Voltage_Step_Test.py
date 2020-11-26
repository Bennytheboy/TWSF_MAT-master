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

for test in range(1, 13):
    case = [13, 16, 19, 22, 16, 22, 16, 22,  1,  4,
             1,  4]

    acc = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.3, 0.3, 1.0, 1.0,
           0.3, 0.3]

    ts = [0.002, 0.002, 0.002, 0.002, 0.001, 0.001, 0.001, 0.001, 0.001, 0.001,
          0.001, 0.001]

    n_iteration = 100  # Number of max. iterations
    tolerance = 0.0001  # Convergence Tolerance
    acceleration = acc[test-1]  # acceleration sfactor used in the network solution
    integration_step = ts[test-1]  # Simulation time step[s]
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
    psspy.addmodellibrary(CasePath + 'SMASC_E161_SMAPPC_E130_342_IVF150.dll')
    psspy.dyre_new([1, 1, 1, 1], CasePath + "Tamworth_SMIB_E161_E130.dyr", '', '', '')
    psspy.branch_data(bus_flt, bus_inf, '1', realar1=0, realar2=0.0001)
    ierr, line_rx = psspy.brndt2(bus_flt, bus_inf, '1', 'RX')
    psspy.ltap(bus_flt, bus_inf, r"""1""", 0.0001, bus_IDTRF, r"""IDTRF""", _f)
    psspy.plant_data(bus_inf, realar1=1.0)
    psspy.purgbrn(bus_IDTRF, bus_flt, r"""1""")
    psspy.two_winding_data_3(bus_flt, bus_IDTRF, r"""1""",
                             [1, bus_flt, 1, 0, 0, 0, 33, 0, bus_flt, 0, 1, 0, 1, 1, 1],
                             [0.0, 0.0001, 100.0, 1.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0, 0.0,
                              0.0, 1.1, 0.9, 1.1, 0.9, 0.0, 0.0, 0.0], r"""IDTRF""")

    OutputFilePath = FigurePath + '07. Grid Voltage Step Test' + "\\" + 'Test' + str(
        test+42) + '_' + FileName + '_sFac' + str(acceleration) + '_dT' + str(
        integration_step) + '_Grid Voltage Step Test.out'
    # ! Setup Dynamic Simulation parameters
    psspy.dynamics_solution_param_2([n_iteration, _i, _i, _i, _i, _i, _i, _i],
                                    [acceleration, tolerance, integration_step, _f, _f, _f, _f, _f])
    psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])

    # convert load , do not change
    psspy.cong(0)
    psspy.conl(0, 1, 1, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.conl(0, 1, 2, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.conl(0, 1, 3, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.ordr(0)  # ! Order the matrix: ORDR
    psspy.fact()  # ! Factorize the matrix: FACT
    psspy.tysl(0)  # ! TYSL

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

    # Run dynamic simulations
    psspy.strt_2([0, 0], OutputFilePath)
    psspy.run(0, 1.0, 1000, 5, 5)
    psspy.two_winding_chng_5(bus_flt, bus_IDTRF, r"""1""", realari4=1.10)
    psspy.run(0, 10.0, 1000, 5, 5)
    psspy.two_winding_chng_5(bus_flt, bus_IDTRF, r"""1""", realari4=0.90)
    psspy.run(0, 20.0, 1000, 5, 5)
    psspy.two_winding_chng_5(bus_flt, bus_IDTRF, r"""1""", realari4=1.00)
    psspy.run(0, 30.0, 1000, 5, 5)