"""
=========================================================================
Python Script
"Generate Passengers"
for use as integrated script with PTV Vissim example "Drop-off Zone"

Copyright (c) Sven Beller, PTV AG. Adapted by Jochen Lohmiller.
All rights reserved.
=========================================================================
"""

import ctypes            # For message box. A library included with Python installation.

ALIGHT_INTERVAL = 2       # [s], time between two exiting passengers
PAX_PEDTYPE = 100         # pedestrian type of pax
PAX_SPEED = 3.6           # [m/s] speed of alighting pax

#==============================================================================================
def Initialization():
    """ General initialization, e.g. assigning values to the global variables """

    # ---------------------------------------------------------------------------------------------
    # Declarations of global variables
    # ---------------------------------------------------------------------------------------------
    global laybyLinkNo               # Vissim link no. of the layby link where the parking lot is placed
    global laybyLink                 # Vissim object of the layby link where the parking lot is placed
    global areaNo                    # Vissim area number where pax should be generated
    global parkingDwellTmDistr       # Number of the dwell time distribution associated with the parking lot where pax exit
    global parkingDwellTmLowerBound  # Lower bound of the dwell time distribution associated with the parking lot
    global currentScriptFileNoPath   # Name of the current script without the path information

    # Get the filename of the script
    pos = str(CurrentScriptFile).rfind('\\')
    currentScriptFileNoPath = str(CurrentScriptFile)[pos+1:]

    # Get the script-associated UDAs
    laybyLinkNo = GetAndCheckScriptUDA("RelLinkNo")
    if laybyLinkNo == 0:
        return

    areaNo = GetAndCheckScriptUDA("MainObjNo")
    if areaNo == 0:
        return

    # Associate the layby link object
    laybyLink = Vissim.Net.Links.ItemByKey(laybyLinkNo)
    laybyLink.SetAttValue("AlightPax", None)    # Reset the label which shows the alighting pax

    # Get the dwell time distribution of the parking lot
    for parkingLotRoute in laybyLink.VehRoutPark:
        # should be called only once because only one parking route should include this link
        parkingDwellTmDistr = parkingLotRoute.VehRoutDec.AttValue("ParkDur(1)")

    if int(parkingDwellTmDistr) <= 0:
        msgString = "Dwell time distribution associated with the parking lot was not found. Default distribution no. 1 chosen."
        ctypes.windll.user32.MessageBoxA(0, msgString, "Warning", 0)
        parkingDwellTmDistr = 1

    # Get lower bound of dwell time
    parkingDwellTmLowerBound = Vissim.Net.TimeDistributions.ItemByKey(parkingDwellTmDistr).AttValue("LowerBound")


#==============================================================================================
def Main():
    """ Main program to be executed during the simulation """
    BoundsCheck()
    GeneratePassengers()

#==============================================================================================
def GeneratePassengers():
    """
    Generates passengers (pax) on area <areaNo> for the car on the <laybyLink>.
    Required globals: laybyLink, ALIGHT_INTERVAL, PAX_PEDTYPE, areaNo, PAX_SPEED
    Required link UDA: AlightPax (to show the number of alighting pax as label)

    The number of pax is car occupancy - 1 (the driver does not exit).
    The first pax exits if the remaining dwell time of the car in the parking lot is less then
    <ALIGHT_INTERVAL> seconds. Each subsequent pax alights <ALIGHT_INTERVAL> seconds later.
    """

    for veh in laybyLink.Vehs:
        if veh.AttValue("DesSpeed") == 0:      # When DesSpeed changes to 0, DwellTime is > 0
            if laybyLink.AttValue("AlightPax") is None:
                laybyLink.SetAttValue("AlightPax", veh.AttValue("Occup") -1)  # show the number of exiting passengers
            if veh.AttValue("Occup") > 1:
                if veh.AttValue("DwellTm") < ALIGHT_INTERVAL:
                    Vissim.Net.Pedestrians.AddPedestrianOnAreaAtCoordinate(PAX_PEDTYPE, areaNo, 0, 0, 0, -1, PAX_SPEED)
                    veh.SetAttValue("Occup", veh.AttValue("Occup") - 1)
                    veh.SetAttValue("DwellTm", veh.AttValue("DwellTm") + ALIGHT_INTERVAL)
        else:
            laybyLink.SetAttValue("AlightPax", None)

#==============================================================================================
# Helpers
#==============================================================================================
def GetAndCheckScriptUDA(udaName):
    """
    Reads the value of script UDA 'UdaName' and returns it if > 0.
    Otherwise stops the simulation and returns 0.
    Required globals: currentScriptFileNoPath.
    """

    if CurrentScript.AttValue(udaName) <= 0:
        msgString = "Please enter a valid number for the script attribute '" + udaName + "'\n" + "for '" + currentScriptFileNoPath+ "'"
        ctypes.windll.user32.MessageBoxA(0, msgString, "Warning", 0)
        Vissim.Simulation.Stop()
        return 0

    return CurrentScript.AttValue(udaName)

#==============================================================================================
def BoundsCheck():
    """
    Ensures that the script period is small enough for the script to run correctly.
    Required globals: parkingDwellTmLowerBound, ALIGHT_INTERVAL, currentScriptFileNoPath

    As the script period may be changed during a simulation run, the bounds check is not only run
    before sim start but every time the script is executed.
    The script GeneratePassengers needs to be exectured at least as often as passengers are allowed
    to alight (ALIGHT_INTERVAL) and as often as the min. dwell time before the first passenger exits.
    """

    simRes = Vissim.Simulation.AttValue("SimRes")

    if parkingDwellTmLowerBound < ALIGHT_INTERVAL:
        maxPeriod = parkingDwellTmLowerBound * simRes
    else:
        maxPeriod = ALIGHT_INTERVAL * simRes

    scriptPeriod = CurrentScript.AttValue("Period")

    if maxPeriod == 0:
        msgString = "You need to increase the simulation resolution in order for the script 'Generate Passengers' to run correctly. The simulation will be stopped now."
        ctypes.windll.user32.MessageBoxA(0, msgString, "Warning", 0)
        Vissim.Simulation.Stop()
    else:
        if scriptPeriod > maxPeriod:
            msgString = "The script period of " + str(scriptPeriod) + " for '" + currentScriptFileNoPath + "'\n"
            msgString += "is too coarse for the currently defined\n"
            msgString += "alighting interval and alighting stop time.\n"
            msgString += "Hence it is reduced to the maximum value of " + str(maxPeriod) + "."
            ctypes.windll.user32.MessageBoxA(0, msgString, "Warning", 0)

            CurrentScript.SetAttValue("Period", maxPeriod)

#==============================================================================================
# End of script
#==============================================================================================
