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
    print("lane count for link ", linkNum, ": ", lane_count)
    return max(lane_count.values())


# ========================================================================
# Simulation set up
# ========================================================================

# Chose Random Seed
Random_Seed = 42
Vissim.Simulation.SetAttValue("RandSeed", Random_Seed)
# Vissim.Simulation.SetAttValue("UseMaxSimSpeed", True)

# Set end of simulation
end_of_simulation = 4800  # simulation second [s]
Vissim.Simulation.SetAttValue("SimPeriod", end_of_simulation)

# step_time = 2
# Vissim.Simulation.SetAttValue("SimRes", step_time)

# ========================================================================
# Signal Timing Plan set up
# ========================================================================

link23 = Vissim.Net.Links.ItemByKey(23)  # between yonge and willowdale
link27 = Vissim.Net.Links.ItemByKey(27)  # between maxome and willowdale

# set alternating green and red times only in yonge intersection
yongeController = Vissim.Net.SignalControllers.ItemByKey(3)
willowdaleController = Vissim.Net.SignalControllers.ItemByKey(4)

# Yonge Signal Groups (Phases)
# 1: NBL, 2: SB, 3: EBL, 4: WB, 5: SBL, 6: NB, 7: WBL, 8: EB

# Willowdale Signal Groups (Phases)
# 2: WB, 4: NB, 5: WBL, 6: EB, 7: NBL, 8: SB

red = "RED"
green = "GREEN"

# make sure there is all red period and ignore ambers

# set yonge phases
yongeNbLeft = yongeController.SGs.ItemByKey(1)
yongeEbLeft = yongeController.SGs.ItemByKey(3)
yongeSbLeft = yongeController.SGs.ItemByKey(5)
yongeWbLeft = yongeController.SGs.ItemByKey(7)

yongeSb = yongeController.SGs.ItemByKey(2)
yongeWb = yongeController.SGs.ItemByKey(4)
yongeNb = yongeController.SGs.ItemByKey(6)
yongeEb = yongeController.SGs.ItemByKey(8)

# set willowdale phases
willowWbLeft = willowdaleController.SGs.ItemByKey(5)
willowNbLeft = willowdaleController.SGs.ItemByKey(7)

willowSb = willowdaleController.SGs.ItemByKey(8)
willowWb = willowdaleController.SGs.ItemByKey(2)
willowNb = willowdaleController.SGs.ItemByKey(4)
willowEb = willowdaleController.SGs.ItemByKey(6)


currentTime = 1
setSimulationBreak(currentTime)

# set everything to red
yongeNbLeft.SetAttValue("SigState", red)
yongeEbLeft.SetAttValue("SigState", red)
yongeSbLeft.SetAttValue("SigState", red)
yongeWbLeft.SetAttValue("SigState", red)

willowWbLeft.SetAttValue("SigState", red)
willowNbLeft.SetAttValue("SigState", red)

willowNb.SetAttValue("SigState", red)
willowSb.SetAttValue("SigState", red)
willowWb.SetAttValue("SigState", red)
willowEb.SetAttValue("SigState", red)

yongeNb.SetAttValue("SigState", red)
yongeSb.SetAttValue("SigState", red)
yongeWb.SetAttValue("SigState", red)
yongeEb.SetAttValue("SigState", red)

networkPerformance = Vissim.Net.VehicleNetworkPerformanceMeasurement

# offset variables
baseTime = 28  # the time it takes for one queued vehicle to go to next intersection
increment = 2  # in a queued platoon, the time it takes for subsequent cars to pass the next intersection
maxQueue = 10  # to calculate new time, do base + (queue length - 1) * increment


offset = 30
cycleLength = 110


while currentTime < end_of_simulation:
    # start with north left turn from willowdale onto steeles
    cycleStart = currentTime

    queue9Length = calculateQueue(9, [1, 2])

    if queue9Length > 0:
        print("willow northbound left green time: ", currentTime)
        willowNbLeft.SetAttValue("SigState", green)

        currentTime += 7  # max northbound
        setSimulationBreak(currentTime)

        print("willow northbound left red time: ", currentTime)
        willowNbLeft.SetAttValue("SigState", red)

    # start of this loop assumes the first intersection is red
    queue23Length = calculateQueue(23, [1, 2])
    print("queue length on link 23: ", queue23Length)

    if queue23Length <= 1:
        offset = 1
    elif queue23Length > 10:
        offset = 9 * increment
    else:
        offset = (queue23Length - 1) * increment

    # change first westbound/eastbound intersection to green
    willowSb.SetAttValue("SigState", red)
    willowNb.SetAttValue("SigState", red)

    print("willowdale green time: ", currentTime)
    # willowGreen = currentTime

    willowWb.SetAttValue("SigState", green)
    willowEb.SetAttValue("SigState", green)

    currentTime += baseTime
    setSimulationBreak(currentTime)

    # change the next intersection to green after baseTime
    yongeSb.SetAttValue("SigState", red)
    yongeNb.SetAttValue("SigState", red)

    print("yonge green time: ", currentTime)
    yongeWb.SetAttValue("SigState", green)
    yongeEb.SetAttValue("SigState", green)

    # now both intersections are green
    # change willowdale to red after offset

    currentTime += offset + 1
    setSimulationBreak(currentTime)

    willowSb.SetAttValue("SigState", green)
    willowNb.SetAttValue("SigState", green)

    print("willow red time: ", currentTime)
    willowWb.SetAttValue("SigState", red)
    willowEb.SetAttValue("SigState", red)

    currentTime += baseTime
    setSimulationBreak(currentTime)

    yongeSb.SetAttValue("SigState", green)
    yongeNb.SetAttValue("SigState", green)

    print("yonge red time: ", currentTime)
    yongeWb.SetAttValue("SigState", red)
    yongeEb.SetAttValue("SigState", red)

    currentTime = cycleStart + cycleLength
    setSimulationBreak(currentTime)


# will stop at the last break, run continuous again to go to end of simulation
Vissim.Simulation.RunContinuous()


# To stop the simulation:
Vissim.Simulation.Stop()

Vissim = None
