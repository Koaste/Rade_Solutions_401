import win32com.client as com
from win32com.client import Dispatch
from pprint import pprint
import os

# Connect to Vissim
Vissim = com.gencache.EnsureDispatch("Vissim.Vissim")
Path_of_COM = "C:\code\CIVE401\Vissim Files\Feb 19 Vissim Run"

# Load a Vissim Network
Filename = os.path.join(Path_of_COM, "Base Model Routed Vehicles.inpx")
flag_read_additionally = False
Vissim.LoadNet(Filename, flag_read_additionally)

# Load Layout:
Filename = os.path.join(Path_of_COM, "Base Model Routed Vehicles.layx")
Vissim.LoadLayout(Filename)


def printSimulationInfo():
    ctime = Vissim.Simulation.AttValue("SimSec")
    print("simulation time: ", ctime)
    # get all vehicles in the work at the actual simulation second
    All_Vehicles = Vissim.Net.Vehicles.GetAll()
    Vissim.Log(16384, "All vehicles by GetAll():")
    print("All vehicles by GetAll():")

    # Link 27 (Steeles between Yonge and Willowdale)
    link_27_count = 0
    # Link 66 (right turn steeles onto yonge)
    link_66_count = 0
    # Link 53 (left turn steeles onto yonge)
    link_53_count = 0
    # Link 23 (Steeles between Willdowdale and Maxome)
    link_23_count = 0

    for cur_Veh in All_Vehicles:
        veh_number = cur_Veh.AttValue("No")
        veh_type = cur_Veh.AttValue("VehType")
        veh_speed = cur_Veh.AttValue("Speed")
        veh_position = cur_Veh.AttValue("Pos")
        veh_linklane = cur_Veh.AttValue("Lane")
        Vissim.Log(
            16384,
            "%s  |  %s  |  %.2f  |  %.2f  |  %s"
            % (veh_number, veh_type, veh_speed, veh_position, veh_linklane),
        )
        # print(
        #     "%s  |  %s  |  %.2f  |  %.2f  |  %s"
        #     % (veh_number, veh_type, veh_speed, veh_position, veh_linklane)
        # )

        split = veh_linklane.split("-")
        link = split[0]

        if int(link) == 27:
            link_27_count += 1
        elif int(link) == 23:
            link_23_count += 1
        elif int(link) == 66:
            link_66_count += 1
        elif int(link) == 53:
            link_53_count += 1

    print("link 23 count ", link_23_count)
    print("link 27 count ", link_27_count)
    print("link 66 count ", link_66_count)
    print("link 53 count ", link_53_count)

    # Get the state of a signal head:
    SH_number = 28  # SH = SignalHead
    State_of_SH = Vissim.Net.SignalHeads.ItemByKey(SH_number).AttValue(
        "SigState"
    )  # possible output see COM Help: SignalizationState Enumeration
    Vissim.Log(
        16384, "Actual state of SignalHead(%d) is: %s" % (SH_number, State_of_SH)
    )
    print("Actual state of SignalHead(%d) is: %s" % (SH_number, State_of_SH))

    SH_number = 75  # SH = SignalHead
    State_of_SH = Vissim.Net.SignalHeads.ItemByKey(SH_number).AttValue(
        "SigState"
    )  # possible output see COM Help: SignalizationState Enumeration
    Vissim.Log(
        16384, "Actual state of SignalHead(%d) is: %s" % (SH_number, State_of_SH)
    )
    print("Actual state of SignalHead(%d) is: %s" % (SH_number, State_of_SH))


def setSimulationBreak(sim_break):
    Vissim.Simulation.SetAttValue("SimBreakAt", sim_break)
    Vissim.Simulation.RunContinuous()


# ========================================================================
# Simulation set up
# ========================================================================

# Chose Random Seed
Random_Seed = 42
Vissim.Simulation.SetAttValue("RandSeed", Random_Seed)
Vissim.Simulation.SetAttValue("UseMaxSimSpeed", True)

# Set end of simulation
end_of_simulation = 4800  # simulation second [s]
Vissim.Simulation.SetAttValue("SimPeriod", end_of_simulation)

# run a single step to set the simulation timings
Vissim.Simulation.RunSingleStep()

# step_time = 2
# Vissim.Simulation.SetAttValue("SimRes", step_time)


# ========================================================================
# Signal Timing Plan set up
# ========================================================================

# set alternating green and red times only in yonge intersection
yongeController = Vissim.Net.SignalControllers.ItemByKey(3)

# Signal Groups (Phases)
# 1: NBL, 2: SB, 3: EBL, 4: WB, 5: SBL, 6: NB, 7: WBL, 8: EB

red = "RED"
green = "GREEN"

# make sure there is all red period and ignore ambers

nbLeft = yongeController.SGs.ItemByKey(1)
ebLeft = yongeController.SGs.ItemByKey(3)
sbLeft = yongeController.SGs.ItemByKey(5)
wbLeft = yongeController.SGs.ItemByKey(7)

# set all the left turns to red
nbLeft.SetAttValue("SigState", red)
ebLeft.SetAttValue("SigState", red)
sbLeft.SetAttValue("SigState", red)
wbLeft.SetAttValue("SigState", red)

# set SB/WB/NB/EB phases
sb = yongeController.SGs.ItemByKey(2)
wb = yongeController.SGs.ItemByKey(4)
nb = yongeController.SGs.ItemByKey(6)
eb = yongeController.SGs.ItemByKey(8)

currentTime = 0
for i in range(end_of_simulation // 60):
    currentTime += 60
    setSimulationBreak(currentTime)

    if i % 2 == 0:
        sb.SetAttValue("SigState", green)
        nb.SetAttValue("SigState", green)

        wb.SetAttValue("SigState", red)
        eb.SetAttValue("SigState", red)
    else:
        sb.SetAttValue("SigState", red)
        nb.SetAttValue("SigState", red)

        wb.SetAttValue("SigState", green)
        eb.SetAttValue("SigState", green)


# will stop at the last break, run continuous again to go to end of simulation
Vissim.Simulation.RunContinuous()


# To stop the simulation:
Vissim.Simulation.Stop()

Vissim = None
