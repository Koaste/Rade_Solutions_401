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
        if link == linkNum and speed < 5 and lane in lane_count:
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

yongeController = Vissim.Net.SignalControllers.ItemByKey(3)
willowdaleController = Vissim.Net.SignalControllers.ItemByKey(4)

# Yonge Signal Groups (Phases)
# 1: NBL, 2: SB, 3: EBL, 4: WB, 5: SBL, 6: NB, 7: WBL, 8: EB

# Willowdale Signal Groups (Phases)
# 2: WB, 4: NB, 5: WBL, 6: EB, 7: NBL, 8: SB

red = "RED"
green = "GREEN"
amber = "AMBER"

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

# timing variables
offset = 30  # the time it takes for one queued vehicle to go to next intersection
increment = 2  # in a queued platoon, the time it takes for subsequent cars to pass the next intersection
maxQueue = 10  # to calculate new time, do base + (queue length - 1) * increment

advanceLeftTime = 6
totalNorthTime = 50
totalWestTime = 60

cycleLength = 110

# queue9Length = calculateQueue(9, [1, 2])

# willowdale timings

willowNBLStart = 2

while True:
    queue9Length = calculateQueue(9, [2])

    willowWestStart = willowNBLStart + 50  # willow north stop, + total north time

    willowWestStop = willowWestStart + 60  # + total west time

    # min offset is 12
    if queue9Length == 0:
        offset = 30
    elif queue9Length > 0:
        offset = 30 - queue9Length * 2
        offset = max(offset, 12)

    yongeNorthStart = willowNBLStart + offset
    yongeWestStart = willowWestStart + offset

    # up date advanceLeftTime here
    setSimulationBreak(willowNBLStart)
    queue9Length = calculateQueue(9, [2])
    if queue9Length > 0:
        advanceLeftTime = 6

        if queue9Length > 1:
            advanceLeftTime = 7
        willowNBLStop = (
            willowNBLStart + advanceLeftTime
        )  # willow north start , + advanceLeftTime

        # start willow north bound left

        willowWb.SetAttValue("SigState", red)
        willowEb.SetAttValue("SigState", red)

        willowNbLeft.SetAttValue("SigState", green)

        # stop willow north bound left and start willow north

        setSimulationBreak(willowNBLStop)
        willowNbLeft.SetAttValue("SigState", amber)

        setSimulationBreak(willowNBLStop + 3)
        willowNbLeft.SetAttValue("SigState", red)

    willowNb.SetAttValue("SigState", green)
    willowSb.SetAttValue("SigState", green)

    # stop yonge west and start yonge north
    setSimulationBreak(yongeNorthStart)
    yongeWb.SetAttValue("SigState", amber)
    yongeEb.SetAttValue("SigState", amber)

    setSimulationBreak(yongeNorthStart + 3)
    yongeWb.SetAttValue("SigState", red)
    yongeEb.SetAttValue("SigState", red)

    queue63Length = calculateQueue(63, [1])

    if queue63Length > 0:
        advanceLeftTime = 6

        if queue9Length > 1:
            advanceLeftTime = 7
        yongeNBLStop = yongeNorthStart + advanceLeftTime

        yongeNbLeft.SetAttValue("SigState", green)

        setSimulationBreak(yongeNBLStop)
        yongeNbLeft.SetAttValue("SigState", amber)

        setSimulationBreak(yongeNBLStop + 3)
        yongeNbLeft.SetAttValue("SigState", red)

    yongeSb.SetAttValue("SigState", green)
    yongeNb.SetAttValue("SigState", green)

    # stop willow north/south and start willow west
    setSimulationBreak(willowWestStart)
    willowNb.SetAttValue("SigState", amber)
    willowSb.SetAttValue("SigState", amber)

    setSimulationBreak(willowWestStart + 3)
    willowNb.SetAttValue("SigState", red)
    willowSb.SetAttValue("SigState", red)

    willowWb.SetAttValue("SigState", green)
    willowEb.SetAttValue("SigState", green)

    # start yonge west and stop yonge north
    setSimulationBreak(yongeWestStart)
    yongeSb.SetAttValue("SigState", amber)
    yongeNb.SetAttValue("SigState", amber)

    setSimulationBreak(yongeWestStart + 3)
    yongeSb.SetAttValue("SigState", red)
    yongeNb.SetAttValue("SigState", red)

    yongeWb.SetAttValue("SigState", green)
    yongeEb.SetAttValue("SigState", green)

    # stop willow west and restart cycle with willow north bound left
    setSimulationBreak(willowWestStop - 3)
    willowWb.SetAttValue("SigState", amber)
    willowEb.SetAttValue("SigState", amber)

    setSimulationBreak(willowWestStop)
    willowWb.SetAttValue("SigState", red)
    willowEb.SetAttValue("SigState", red)

    willowNBLStart = willowWestStop + 1


# will stop at the last break, run continuous again to go to end of simulation
Vissim.Simulation.RunContinuous()

# To stop the simulation:
Vissim.Simulation.Stop()

Vissim = None
