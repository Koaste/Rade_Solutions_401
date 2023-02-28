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


def calculateQueue(linkNum, lanes):
    vehicles = Vissim.Net.Vehicles.GetAll()
    lane_count = {}

    for lane in lanes:
        lane_count[lane] = 0

    for vehicle in vehicles:
        raw = vehicle.AttValue("Lane")
        link_lane_arr = raw.split("-")
        link = int(link_lane_arr[0])
        lane = int(link_lane_arr[1])

        speed = float(vehicle.AttValue("Speed"))

        # a vehicle is considered queued if its speed is less than 5km/h
        if link == linkNum and speed < 5:
            lane_count[lane] += 1
    print(lane_count)
    return max(lane_count.values())


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
    printSimulationInfo()


# ========================================================================
# Simulation
# ========================================================================

# Chose Random Seed
Random_Seed = 42
Vissim.Simulation.SetAttValue("RandSeed", Random_Seed)
Vissim.Simulation.SetAttValue("UseMaxSimSpeed", True)

# Set end of simulation
end_of_simulation = 4800  # simulation second [s]
Vissim.Simulation.SetAttValue("SimPeriod", end_of_simulation)

step_time = 3
Vissim.Simulation.SetAttValue("SimRes", step_time)

setSimulationBreak(1)

# set the state of a signal controller:
SC_number = 3  # SC = SignalController

SignalController = Vissim.Net.SignalControllers.ItemByKey(SC_number)
willow = Vissim.Net.SignalControllers.ItemByKey(4)

# new_state = "RED"
# for i in range(1, 9):
#     signalGroup = SignalController.SGs.ItemByKey(i)
#     signalGroup.SetAttValue("SigState", "GREEN")

# for i in range(2, 9):
#     if i == 3:
#         continue
#     signalGroup = willow.SGs.ItemByKey(i)
#     signalGroup.SetAttValue("SigState", "RED")

# setSimulationBreak(150)

# print(calculateQueue(23, [1, 2]))

# for i in range(2, 9):
#     if i == 3:
#         continue
#     signalGroup = willow.SGs.ItemByKey(i)
#     signalGroup.SetAttValue("SigState", "GREEN")

# datapoints = Vissim.Net.VehicleTravelTimeMeasurements
# datapoint1 = datapoints.ItemByKey(1)
# datapoint2 = datapoints.ItemByKey(2)


# datapoint5 = datapoints.ItemByKey(3)
# print("test")

# datapoint6 = datapoints.ItemByKey(4)

# print(datapoint1.AttValue("TravTm(Current, Last, All)"))
# print(datapoint2.AttValue("TravTm(Current, Last, All)"))


# will stop at the last break, run continuous again to go to end of simulation
Vissim.Simulation.RunContinuous()


# To stop the simulation:
Vissim.Simulation.Stop()

Vissim = None
