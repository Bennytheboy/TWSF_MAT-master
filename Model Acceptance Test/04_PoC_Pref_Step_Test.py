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

for test in range(1, 19):
    case = [13, 14, 15, 16, 17, 18, 16, 17, 18, 16,
            17, 18, 1, 2, 3, 1, 2, 3]

    acc = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.3,
           0.3, 0.3, 1.0, 1.0, 1.0, 0.3, 0.3, 0.3]

    ts = [0.002, 0.002, 0.002, 0.002, 0.002, 0.002, 0.001, 0.001, 0.001, 0.002,
          0.002, 0.002, 0.001, 0.001, 0.001, 0.002, 0.002, 0.002]

    n_iteration = 100  # Number of max. iterations
    tolerance = 0.0001  # Convergence Tolerance
    acceleration = acc[test-1]  # acceleration sfactor used in the network solution
    integration_step = ts[test-1]  # Simulation time step[s]

    _i = psspy.getdefaultint()
    _f = psspy.getdefaultreal()
    _s = psspy.getdefaultchar()
    redirect.psse2py()
    psspy.psseinit(50000)

    FileName = f_list[case[test - 1] - 1]
    psspy.case(CasePath + FileName)
    psspy.resq(CasePath + "Tamworth_SMIB.seq")
    psspy.addmodellibrary(CasePath + 'SMASC_E161_SMAPPC_E130_342_IVF150.dll')
    psspy.dyre_new([1, 1, 1, 1], CasePath + "Tamworth_SMIB_E161_E130.dyr", '', '', '')

    dPref = 0.2
    MBASE = 83.6
    CON_PPCPref = 18
    VAR_PPCPini = 2

    if test < 10:
        OutputFilePath = FigurePath + '04. PoC Pref Step Test' + "\\" + 'Test0' + str(
            test) + '_' + FileName + '_sFac' + str(acceleration) + '_dT' + str(
            integration_step) + '_PoC Pref Step Test.out'
    else:
        OutputFilePath = FigurePath + '04. PoC Pref Step Test' + "\\" + 'Test' + str(
            test) + '_' + FileName + '_sFac' + str(acceleration) + '_dT' + str(
            integration_step) + '_PoC Pref Step Test.out'
    # ! Setup Dynamic Simulation parameters
    psspy.dynamics_solution_param_2([n_iteration, _i, _i, _i, _i, _i, _i, _i],
                                    [acceleration, tolerance, integration_step, _f, _f, _f, _f, _f])
    # psspy.change_plmod_con(100, r"""1""", r"""SMAPPC127""", 16, 0.2)
    psspy.fdns([0, 0, 1, 1, 0, 0, 99, 0])

    # convert load , do not change
    psspy.cong(0)
    psspy.conl(0, 1, 1, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.conl(0, 1, 2, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.conl(0, 1, 3, [0, 0], [100.0, 0.0, 0.0, 100.0])
    psspy.ordr(0)  # Order the matrix: ORDR
    psspy.fact()  # Factorize the matrix: FACT
    psspy.tysl(0)  # TYSL

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

    psspy.strt_2([0, 0], OutputFilePath)
    J_vals = []
    L_vals = []
    Pset_vals = []
    ierr, J1 = psspy.mdlind(100, '1', 'EXC', 'CON')
    ierr, L1 = psspy.mdlind(100, '1', 'EXC', 'VAR')
    ierr, Pset1 = psspy.dsrval('VAR', L1 + VAR_PPCPini)
    J_vals.append(J1)
    L_vals.append(L1)
    Pset_vals.append(Pset1)

    # Run dynamic simulations
    psspy.run(0, 5.0, 1000, 5, 5)
    psspy.change_plmod_con(100, r"""1""", r"""SMAPPC130""", 19, Pset_vals[0] * (1 - dPref))
    psspy.run(0, 10.0, 1000, 5, 5)
    psspy.change_plmod_con(100, r"""1""", r"""SMAPPC130""", 19, Pset_vals[0] * (1 - dPref*2))
    psspy.run(0, 15.0, 1000, 5, 5)
    psspy.change_plmod_con(100, r"""1""", r"""SMAPPC130""", 19, Pset_vals[0] * (1 - dPref*3))
    psspy.run(0, 20.0, 1000, 5, 5)
    psspy.change_plmod_con(100, r"""1""", r"""SMAPPC130""", 19, Pset_vals[0] * (1 - dPref*4))
    psspy.run(0, 25.0, 1000, 5, 5)
    psspy.change_plmod_con(100, r"""1""", r"""SMAPPC130""", 19, Pset_vals[0])
    psspy.run(0, 40.0, 1000, 5, 5)