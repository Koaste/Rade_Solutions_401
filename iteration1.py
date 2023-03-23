import win32com.client as com
from win32com.client import Dispatch
from pprint import pprint
import os

# connect to Vissim
Vissim = com.gencache.EnsureDispatch("Vissim.Vissim")
Path_of_COM = "C:\code\CIVE401\Vissim Files\Feb 27 Vissim Run"

# load a Vissim Network
Filename = os.path.join(Path_of_COM, "Feb 27 Base Model Routed.inpx")
flag_read_additionally = False
Vissim.LoadNet(Filename, flag_read_additionally)

# load layout:
Filename = os.path.join(Path_of_COM, "Feb 27 Base Model Routed.layx")
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


def initializeMovements(controller, validMovements):
    movementDict = {
        1: "WBLeft",
        2: "EB",
        3: "SBLeft",
        4: "NB",
        5: "EBLeft",
        6: "WB",
        7: "NBLeft",
        8: "SB",
    }

    ret = {}

    for movement in validMovements:
        ret[movementDict[movement]] = controller.SGs.ItemByKey(movement)

    return ret


# ========================================================================
# Simulation set up
# ========================================================================

# choose random simulation eed
Random_Seed = 42
Vissim.Simulation.SetAttValue("RandSeed", Random_Seed)

# set end of simulation
end_of_simulation = 4800  # simulation second [s]
Vissim.Simulation.SetAttValue("SimPeriod", end_of_simulation)

yongeController = Vissim.Net.SignalControllers.ItemByKey(3)
willowdaleController = Vissim.Net.SignalControllers.ItemByKey(4)

# initialize constants
red = "RED"
green = "GREEN"
amber = "AMBER"

# initialize controllers at intersections of focus
controllers = {
    "willowdale": Vissim.Net.SignalControllers.ItemByKey(4),
    "yonge": Vissim.Net.SignalControllers.ItemByKey(3),
    "maxome": Vissim.Net.SignalControllers.ItemByKey(5),
}

# populate hash map with movements for each controller
yongeMovements = initializeMovements(controllers["yonge"], [1, 2, 3, 4, 5, 6, 7, 8])
willowMovements = initializeMovements(controllers["willowdale"], [1, 2, 4, 6, 7, 8])
maxomeMovements = initializeMovements(controllers["maxome"], [2, 4, 6, 8])

currentTime = 1
setSimulationBreak(currentTime)

# set all lights to red for simulation start
for movement in yongeMovements.values():
    movement.SetAttValue("SigState", red)

for movement in willowMovements.values():
    movement.SetAttValue("SigState", red)

# base yonge/willowdale offset value is 30 seconds
yongeWillowOffset = 30

# base maxome/willowdale offset value is 50 seconds
maxomeWillowOffset = 50

# in a queued platoon, the time it takes for subsequent cars to pass the next intersection
increment = 2

# total time given to north/south phase
totalNorthTime = 50

# total time given to east/west phase
totalWestTime = 60

# advance left time initialized to 7 seconds
advanceLeftTime = 7

# one cycle is 110 seconds
cycleLength = 110

willowNBLStart = 2

while True:
    queue9Length = calculateQueue(9, [2])

    willowWestStart = willowNBLStart + 50

    willowWestStop = willowWestStart + 60

    maxomeNorthStart = willowNBLStart + 60

    if queue9Length == 0:
        yongeWillowOffset = 30
    elif queue9Length > 0:
        yongeWillowOffset = 30 - queue9Length * increment
        yongeWillowOffset = max(yongeWillowOffset, 15)

    yongeNorthStart = willowNBLStart + yongeWillowOffset
    yongeWestStart = willowWestStart + yongeWillowOffset

    setSimulationBreak(willowNBLStart)

    maxomeMovements["WB"].SetAttValue("SigState", green)
    maxomeMovements["EB"].SetAttValue("SigState", green)

    queue9Length = calculateQueue(9, [2])
    if queue9Length > 0:
        willowNBLStop = willowNBLStart + advanceLeftTime

        # stop willowdale west/east phases and start willowdale NBL phase
        willowMovements["WB"].SetAttValue("SigState", red)
        willowMovements["EB"].SetAttValue("SigState", red)

        willowMovements["NBLeft"].SetAttValue("SigState", green)

        # stop willow NBL phase and start willowdale north/south phase
        setSimulationBreak(willowNBLStop)
        willowMovements["NBLeft"].SetAttValue("SigState", amber)

        setSimulationBreak(willowNBLStop + 3)
        willowMovements["NBLeft"].SetAttValue("SigState", red)

    willowMovements["NB"].SetAttValue("SigState", green)
    willowMovements["SB"].SetAttValue("SigState", green)

    # stop yonge west/east phase and start yonge north/south phase
    setSimulationBreak(yongeNorthStart)
    yongeMovements["WB"].SetAttValue("SigState", amber)
    yongeMovements["EB"].SetAttValue("SigState", amber)

    setSimulationBreak(yongeNorthStart + 3)
    yongeMovements["WB"].SetAttValue("SigState", red)
    yongeMovements["EB"].SetAttValue("SigState", red)

    # calculate length of yonge NBL queue
    queue63Length = calculateQueue(63, [1])

    # calcualte length of yonge SBL queue
    queue64Length = calculateQueue(64, [1])

    # yonge NBL and SBL advance green calculation
    yongeNBLStop = yongeNorthStart + advanceLeftTime

    # if there are cars in both NBL/SBL queues
    if queue63Length > 0 and queue64Length > 0:
        yongeMovements["NBLeft"].SetAttValue("SigState", green)
        yongeMovements["SBLeft"].SetAttValue("SigState", green)

        setSimulationBreak(yongeNBLStop)
        yongeMovements["NBLeft"].SetAttValue("SigState", amber)
        yongeMovements["SBLeft"].SetAttValue("SigState", amber)

        setSimulationBreak(yongeNBLStop + 3)
        yongeMovements["NBLeft"].SetAttValue("SigState", red)
        yongeMovements["SBLeft"].SetAttValue("SigState", red)
    # only yonge NBL has a queue
    elif queue63Length > 0:
        yongeMovements["NBLeft"].SetAttValue("SigState", green)

        setSimulationBreak(yongeNBLStop)
        yongeMovements["NBLeft"].SetAttValue("SigState", amber)

        setSimulationBreak(yongeNBLStop + 3)
        yongeMovements["NBLeft"].SetAttValue("SigState", red)
    # only yonge SBL has a queue
    elif queue64Length > 0:
        yongeMovements["SBLeft"].SetAttValue("SigState", green)

        setSimulationBreak(yongeNBLStop)
        yongeMovements["SBLeft"].SetAttValue("SigState", amber)

        setSimulationBreak(yongeNBLStop + 3)
        yongeMovements["SBLeft"].SetAttValue("SigState", red)

    # start yonge north/south phase
    yongeMovements["SB"].SetAttValue("SigState", green)
    yongeMovements["NB"].SetAttValue("SigState", green)

    # stop willowdale north/south phase and start willowdale west/east phase
    setSimulationBreak(willowWestStart)

    willowMovements["NB"].SetAttValue("SigState", amber)
    willowMovements["SB"].SetAttValue("SigState", amber)

    setSimulationBreak(willowWestStart + 3)
    willowMovements["NB"].SetAttValue("SigState", red)
    willowMovements["SB"].SetAttValue("SigState", red)

    # calculate willowdale NBL queue
    queue52Length = calculateQueue(52, [1])

    willowMovements["WB"].SetAttValue("SigState", green)

    # check if there is a queue on willowdale NBL
    if queue52Length > 0:
        willowWBLStop = willowWestStart + 3 + 7

        willowMovements["WBLeft"].SetAttValue("SigState", green)

        # if advance left is triggered, maxome north needs to be green for coordination
        setSimulationBreak(maxomeNorthStart - 3)

        maxomeMovements["WB"].SetAttValue("SigState", amber)
        maxomeMovements["EB"].SetAttValue("SigState", amber)

        setSimulationBreak(maxomeNorthStart)

        maxomeMovements["WB"].SetAttValue("SigState", red)
        maxomeMovements["EB"].SetAttValue("SigState", red)

        maxomeMovements["NB"].SetAttValue("SigState", green)
        maxomeMovements["SB"].SetAttValue("SigState", green)

        willowMovements["WBLeft"].SetAttValue("SigState", amber)

        setSimulationBreak(willowWBLStop + 3)
        willowMovements["WBLeft"].SetAttValue("SigState", red)
    else:
        setSimulationBreak(maxomeNorthStart - 3)

        maxomeMovements["WB"].SetAttValue("SigState", amber)
        maxomeMovements["EB"].SetAttValue("SigState", amber)

        setSimulationBreak(maxomeNorthStart)

        maxomeMovements["WB"].SetAttValue("SigState", red)
        maxomeMovements["EB"].SetAttValue("SigState", red)

        maxomeMovements["NB"].SetAttValue("SigState", green)
        maxomeMovements["SB"].SetAttValue("SigState", green)

    willowMovements["EB"].SetAttValue("SigState", green)

    # start yonge west/east phase and stop yonge north/south phase
    setSimulationBreak(yongeWestStart)
    yongeMovements["SB"].SetAttValue("SigState", amber)
    yongeMovements["NB"].SetAttValue("SigState", amber)

    setSimulationBreak(yongeWestStart + 3)
    yongeMovements["SB"].SetAttValue("SigState", red)
    yongeMovements["NB"].SetAttValue("SigState", red)

    yongeMovements["WB"].SetAttValue("SigState", green)
    yongeMovements["EB"].SetAttValue("SigState", green)

    # stop willowdale west/east phase and restart cycle with willowdale NBL phase
    setSimulationBreak(willowWestStop - 3)
    willowMovements["WB"].SetAttValue("SigState", amber)
    willowMovements["EB"].SetAttValue("SigState", amber)

    maxomeMovements["NB"].SetAttValue("SigState", amber)
    maxomeMovements["SB"].SetAttValue("SigState", amber)

    setSimulationBreak(willowWestStop)
    willowMovements["WB"].SetAttValue("SigState", red)
    willowMovements["EB"].SetAttValue("SigState", red)

    maxomeMovements["NB"].SetAttValue("SigState", red)
    maxomeMovements["SB"].SetAttValue("SigState", red)

    willowNBLStart = willowWestStop + 1


# will stop at the last break, run continuous again to go to end of simulation
Vissim.Simulation.RunContinuous()

# To stop the simulation:
Vissim.Simulation.Stop()

Vissim = None
