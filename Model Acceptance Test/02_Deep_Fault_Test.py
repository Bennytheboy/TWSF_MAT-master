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

# Set Simulation Path.
ProgramPath = WorkingFolder + '\\'
CasePath = WorkingFolder + "\\" + 'Cases' + "\\"
FigurePath = WorkingFolder + "\\" + 'Results' + "\\"
f_list = os.listdir(CasePath)

for test in range(1, 31):
    case = [ 1,  2,  3,  4,  5,  6 , 7,  8,  9, 10,
            11, 12, 13, 14, 15, 16, 17, 18,  4,  5,
             6, 10, 11, 12,  4,  5,  6, 10, 11, 12]

    acc = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0,
           1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0,
           1.0, 1.0, 1.0, 1.0, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3]

    ts = [0.002, 0.002, 0.002, 0.002, 0.002, 0.002, 0.002, 0.002, 0.002, 0.002,
          0.002, 0.002, 0.002, 0.002, 0.002, 0.002, 0.002, 0.002, 0.001, 0.001,
          0.001, 0.001, 0.001, 0.001, 0.002, 0.002, 0.002, 0.002, 0.002, 0.002]

    n_iteration = 100  # Number of max. iterations
    tolerance = 0.0001  # Convergence Tolerance
    acceleration = acc[test-1]  # acceleration sfactor used in the network solution
    integration_step = ts[test-1]  # Simulation time step[s]
    Vflt = 0.01
    Tflt = 0.12

    _i = psspy.getdefaultint()
    _f = psspy.getdefaultreal()
    _s = psspy.getdefaultchar()
    redirect.psse2py()
    psspy.psseinit(50000)

    FileName = f_list[case[test - 1]-1]
    psspy.case(CasePath + FileName)
    psspy.resq(CasePath + "Tamworth_SMIB.seq")
    # psspy.addmodellibrary(CasePath + 'SMASC_E161_SMAPPC_E130_344_IVF150.dll')
    psspy.addmodellibrary(CasePath + 'SMASC_E161_SMAPPC_E130_342_IVF150.dll')
    psspy.dyre_new([1, 1, 1, 1], CasePath + "Tamworth_SMIB_E161_E130.dyr", '', '', '')

    if test < 10:
        OutputFilePath = FigurePath + '02. Deep Fault Test' + "\\" + 'Test0' + str(
            test) + '_' + FileName + '_Vflt0.01_Tflt0.12' + '_sFac' + str(acceleration) + '_dT' + str(
            integration_step) + '.out'
    else:
        OutputFilePath = FigurePath + '02. Deep Fault Test' + "\\" + 'Test' + str(
            test) + '_' + FileName + '_Vflt0.01_Tflt0.12' + '_sFac' + str(acceleration) + '_dT' + str(
            integration_step) + '.out'
    # ! Setup Dynamic Simulation parameters
    psspy.dynamics_solution_param_2([n_iteration, _i, _i, _i, _i, _i, _i, _i],
                                    [acceleration, tolerance, integration_step, _f, _f, _f, _f, _f])
    psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])

    # Calculate fault admittance needed to create specified voltage dip
    ierr, Zsys1 = psspy.brndt2(107, 969, '1', 'RX')
    fault_B = ((1 - Vflt) / (Vflt * Zsys1)).imag * 100
    fault_G = ((1 - Vflt) / (Vflt * Zsys1)).real

    # convert load , do not change
    psspy.cong(0)
    psspy.conl(0, 1, 1, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.conl(0, 1, 2, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.conl(0, 1, 3, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.ordr(0) # Order the matrix: ORDR
    psspy.fact()  # Factorize the matrix: FACT
    psspy.tysl(0) # TYSL

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

    [ierr, var_ppc_conp] = psspy.mdlind(100, '1', 'EXC', 'CON')
    [ierr, var_ppc_setp] = psspy.mdlind(100, '1', 'EXC', 'VAR')
    [ierr, var_ppc_mode] = psspy.mdlind(100, '1', 'EXC', 'ICON')
    [ierr, var_inv_con] = psspy.mdlind(100, '1', 'GEN', 'CON')
    [ierr, var_inv_var] = psspy.mdlind(100, '1', 'GEN', 'VAR')
    [ierr, var_inv_mod] = psspy.mdlind(100, '1', 'GEN', 'ICON')

    # Run dynamic simulations
    psspy.strt_2([0, 0], OutputFilePath)
    psspy.run(0, 1.0, 1000, 5, 5)
    psspy.dist_bus_fault(107, 1, 132.0, [fault_G, fault_B])
    psspy.run(0, (1.0 + Tflt), 1000, 5, 5)
    psspy.dist_clear_fault(1)
    psspy.run(0, 10, 1000, 5, 5)