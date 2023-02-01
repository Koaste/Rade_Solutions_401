"""
=========================================================================
Python Script
"Move Barrier"
for use as integrated script with PTV Vissim example "Drop-off Zone"

Copyright (c) Sven Beller, PTV AG. Adapted by Jochen Lohmiller.
All rights reserved.
=========================================================================
"""

import ctypes          # For message box. A library included with Python installation.

BARRIER_ALLOWANCE = 2   # Duration [s] between start of barrier opening until vehicle starts

#==============================================================================================
def Initialization():
    """
    General initialization, e.g. assigning values to the global variables
    """
    global maxState                   # Max. state number for static 3D model of the barrier
    global barrierLinkNo              # Number of the Vissim link/connector next to the barrier
    global barrierObjNo               # Number of the static 3D model that represents the barrier
    global barrier                    # The 3D model that represents the barrier
    global barrierDwellTmDistr        # Number of the dwell time distribution associated with the stop sign at the barrier
    global barrierDwellTmLowerBound   # Lower bound of the dwell time distribution associated with the barrier stop sign
    global currentScriptFileNoPath    # Name of the current script without the path information

    # Get the filename of the script
    pos = str(CurrentScriptFile).rfind('\\')
    currentScriptFileNoPath = str(CurrentScriptFile)[pos+1:]

    # Get the script-associated UDAs
    barrierLinkNo = GetAndCheckScriptUDA("RelLinkNo")
    if barrierLinkNo == 0:
        return

    barrierObjNo = GetAndCheckScriptUDA("MainObjNo")
    if barrierObjNo == 0:
        return

    # Associate the barrier object
    barrier = Vissim.Net.Static3DModels.ItemByKey(barrierObjNo)
    maxState = barrier.AttValue("NumStates") - 1    # first state is 0 (zero-based)
    barrier.SetAttValue("State", 0)

    # Get the dwell time distribution of the stop sign
    for stopSign in Vissim.Net.StopSigns:
        if stopSign.AttValue("Lane\\Link\\No") == barrierLinkNo:
            barrierDwellTmDistr = stopSign.AttValue("DwellTmDistr(10)")

    if int(barrierDwellTmDistr) <= 0:
        msgString = "Dwell time distribution associated with the parking lot was not found. Default distribution no. 1 chosen."
        ctypes.windll.user32.MessageBoxA(0, msgString, "Warning", 0)
        barrierDwellTmDistr = 1

    # Get lower bound of dwell time
    barrierDwellTmLowerBound = Vissim.Net.TimeDistributions.ItemByKey(barrierDwellTmDistr).AttValue("LowerBound")


#==============================================================================================
def Main():
    """
    Main program to be executed during the simulation
    """
    BoundsCheck()
    MoveBarrier()

# ==============================================================================================
def MoveBarrier():
    """
    Controls the barrier opening and closing process.
    Required globals: barrierLinkNo, barrier, maxState, BARRIER_ALLOWANCE

    The barrier starts opening as soon as the vehicle in front of the barrier has a remaining
    dwell time of <BARRIER_ALLOWANCE> seconds or less.
    The barrier closes as soon as the vehicle has left the barrier link.
    It is assumed that never more than one vehicle is located on the barrier link.
    """
    # open barrier
    for veh in Vissim.Net.Links.ItemByKey(barrierLinkNo).Vehs:
        # open the barrier as soon as the dwell time of the vehicle is BARRIER_ALLOWANCE sec or less
        if 0 < int(veh.AttValue("DwellTm")) <= BARRIER_ALLOWANCE  and  barrier.AttValue("State") < maxState:
            barrier.SetAttValue("State", barrier.AttValue("State") + 1)

    # close barrier
    if Vissim.Net.Links.ItemByKey(barrierLinkNo).Vehs.Count == 0 and barrier.AttValue("State") > 0:
        barrier.SetAttValue("State", barrier.AttValue("State") - 1)


#==============================================================================================
# Helpers
#==============================================================================================
def GetAndCheckScriptUDA(udaName):
    """
    Reads the value of script UDA 'udaName' and returns it if > 0.
    Otherwise stops the simulation and returns 0.
    Required globals: currentScriptFileNoPath.
    """
    if int(CurrentScript.AttValue(udaName)) <= 0:
        msgString = "Please enter a valid number for the script attribute '" + udaName + "'\n" + "for '" + currentScriptFileNoPath+ "'"
        ctypes.windll.user32.MessageBoxA(0, msgString, "Warning", 0)
        Vissim.Simulation.Stop()
        return 0

    return CurrentScript.AttValue(udaName)

#==============================================================================================
def BoundsCheck():
    """
    Ensures that the script period is small enough for the script to run correctly.
    Required globals: barrierDwellTmLowerBound, BARRIER_ALLOWANCE, maxState, currentScriptFileNoPath

    As the script period may be changed during a simulation run, the bounds check is not only run
    before sim start but every time the script is executed.
    The script MoveBarrier needs to be executed at least as often as it needs to fully open the barrier
    (<maxState> times (= state switches) within BARRIER_ALLOWANCE seconds)
    """
    simRes = Vissim.Simulation.AttValue("SimRes")

    if barrierDwellTmLowerBound < BARRIER_ALLOWANCE:
        maxPeriod = barrierDwellTmLowerBound * simRes / maxState
    else:
        maxPeriod = BARRIER_ALLOWANCE * simRes / maxState

    scriptPeriod = CurrentScript.AttValue("Period")

    if maxPeriod == 0:
        msgString = "You need to increase the simulation resolution in order for the script 'Move Barrier' to run correctly. The simulation will be stopped now."
        ctypes.windll.user32.MessageBoxA(0, msgString, "Warning", 0)
        Vissim.Simulation.Stop()
    elif scriptPeriod > maxPeriod:
        msgString = "The script period of " + str(scriptPeriod) + " for '" + currentScriptFileNoPath + "'\n"
        msgString += "is too coarse for the currently defined\n"
        msgString += "stopping time and barrier movement.\n"
        msgString += "Hence it is reduced to the maximum value of " + str(maxPeriod) + "."
        ctypes.windll.user32.MessageBoxA(0, msgString, "Warning", 0)

        CurrentScript.SetAttValue("Period", maxPeriod)

#==============================================================================================
# End of script
#==============================================================================================
