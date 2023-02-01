'==========================================================================
' VB Script 
' "Generate Passengers"
' for use as integrated script with PTV Vissim example "Drop-off Zone"
'
' Copyright (c) Sven Beller, PTV AG.
' All rights reserved.
'==========================================================================

Option Explicit

' ---------------------------------------------------------------------------------------------
' Constants
' ---------------------------------------------------------------------------------------------
Const ALIGHT_INTERVAL = 2     ' [s], time between two exiting passengers
Const PAX_PEDTYPE = 100       ' pedestrian type of pax
Const PAX_SPEED = 3.6         ' [m/s] speed of alighting pax

' ---------------------------------------------------------------------------------------------
' Declarations of global variables
' ---------------------------------------------------------------------------------------------
Dim laybyLinkNo               ' Vissim link no. of the layby link where the parking lot is placed
Dim laybyLink                 ' Vissim object of the layby link where the parking lot is placed
Dim areaNo                    ' Vissim area number where pax should be generated
Dim parkingDwellTmDistr       ' Number of the dwell time distribution associated with the parking lot where pax exit
Dim parkingDwellTmLowerBound  ' Lower bound of the dwell time distribution associated with the parking lot
Dim currentScriptFileNoPath   ' name of the current script without the path information

' ---------------------------------------------------------------------------------------------
' Initialization
' ---------------------------------------------------------------------------------------------
Call Initialization   ' is called only once when the script is run for the first time and not thereafter


'==============================================================================================  
Sub Initialization
' General initialization, e.g. assigning values to the global variables
' Required globals: all globally declared variables
  
  ' Get the filename of the script
  dim pos
  pos = InStrRev(CurrentScriptFile, "\")
  currentScriptFileNoPath = Mid(CurrentScriptFile, pos+1)

  ' Get the script-associated UDAs
  laybyLinkNo = GetAndCheckScriptUDA ("RelLinkNo")
  if laybyLinkNo = 0 then exit sub
  
  areaNo = GetAndCheckScriptUDA ("MainObjNo")
  if areaNo = 0 then exit sub

  ' Associate the layby link object
  Set laybyLink = vissim.Net.Links.ItemByKey(laybyLinkNo)    
  laybyLink.AttValue("AlightPax") = ""    ' Reset the label which shows the alighting pax

  ' Get the dwell time distribution of the parking lot
  Dim parkingLotRoute
  for each parkingLotRoute in laybyLink.VehRoutPark
    ' should be called only once because only one parking route should include this link
    parkingDwellTmDistr = parkingLotRoute.VehRoutDec.AttValue("ParkDur(1)")
  next

  if parkingDwellTmDistr <= 0 then 
    msgbox "Dwell time distribution associated with the parking lot was not found. Default distribution no. 1 chosen.", vbExclamation
    parkingDwellTmDistr = 1
  end if

  ' Get lower bound of dwell time
  parkingDwellTmLowerBound = vissim.Net.TimeDistributions.ItemByKey(parkingDwellTmDistr).AttValue("LowerBound")
  
End Sub


'==============================================================================================  
'==============================================================================================  
'==============================================================================================  
Sub Main()
' Main program to be executed during the simulation

	Call BoundsCheck()
	Call GeneratePassengers

End Sub

'==============================================================================================  
Sub GeneratePassengers()  
' Generates passengers (pax) on area <areaNo> for the car on the <laybyLink>.
' Required globals: laybyLink, ALIGHT_INTERVAL, PAX_PEDTYPE, areaNo, PAX_SPEED
' Required link UDA: AlightPax (to show the number of alighting pax as label)

' The number of pax is car occupancy - 1 (the driver does not exit). 
' The first pax exits if the remaining dwell time of the car in the parking lot is less then
' <ALIGHT_INTERVAL> seconds. Each subsequent pax alights <ALIGHT_INTERVAL> seconds later. 

  Dim veh
  For Each veh In laybyLink.Vehs
    If veh.AttValue("DesSpeed") = 0 Then    ' When DesSpeed changes to 0, DwellTime is > 0
        If laybyLink.AttValue("AlightPax") = "" Then
            laybyLink.AttValue("AlightPax") = veh.AttValue("Occup") - 1 'show number of exiting passengers
        End If
        If veh.AttValue("Occup") > 1 Then   ' 1 = driver remains in vehicle
            If (veh.AttValue("DwellTm") < ALIGHT_INTERVAL) Then
                Call vissim.Net.Pedestrians _
                   .AddPedestrianOnAreaAtCoordinate(PAX_PEDTYPE, areaNo, 0, 0, 0, -1, PAX_SPEED)
                veh.AttValue("Occup") = veh.AttValue("Occup") - 1
                veh.AttValue("DwellTm") = veh.AttValue("DwellTm") + ALIGHT_INTERVAL
            End If
        End If
    Else
        laybyLink.AttValue("AlightPax") = ""
    End If
  Next
End Sub

' ==============================================================================================  
' Helpers
' ==============================================================================================  
Function GetAndCheckScriptUDA (udaName)
' Reads the value of script UDA 'UdaName' and returns it if > 0. 
' Otherwise stops the simulation and returns 0. """
' Required globals: currentScriptFileNoPath
 
  if CurrentScript.AttValue(udaName) <= 0 then
    msgbox "Please enter a valid number for the script attribute '" + udaName + "'" + vbcr _
           + "for '" + currentScriptFileNoPath + "'", vbCritical
    Vissim.Simulation.Stop
    GetAndCheckScriptUDA = 0
  else
    GetAndCheckScriptUDA = CurrentScript.AttValue(udaName)
  end if

End Function

'==============================================================================================  
Sub BoundsCheck()
' Ensures that the script period is small enough for the script to run correctly.
' Required globals: parkingDwellTmLowerBound, ALIGHT_INTERVAL, currentScriptFileNoPath

' As the script period may be changed during a simulation run, the bounds check is not only run 
' before sim start but every time the script is executed. 
' The script GeneratePassengers needs to be exectured at least as often as passengers are allowed 
' to alight (ALIGHT_INTERVAL) and as often as the min. dwell time before the first passenger exits.

  Dim simRes
  simRes = vissim.simulation.AttValue("SimRes")

  Dim maxPeriod
  if parkingDwellTmLowerBound < ALIGHT_INTERVAL then 
    maxPeriod = parkingDwellTmLowerBound * simRes
  else
    maxPeriod = ALIGHT_INTERVAL * simRes    
  end if
  
  Dim scriptPeriod
  scriptPeriod = CurrentScript.AttValue("Period")

  If maxPeriod = 0 then
    msgbox "You need to increase the simulation resolution in order for the script 'Generate Passengers' to run correctly. The simulation will be stopped now.", vbCritical
    vissim.simulation.stop
  ElseIf (scriptPeriod > maxPeriod) then 
    msgbox "The script period of " + CStr(scriptPeriod) + " for '" + currentScriptFileNoPath + "'" _ 
           + vbCrLf + "is too coarse for the currently defined" _
           + vbCrLf + "alighting interval and alighting stop time." _ 
           + vbCrLf + "Hence it is reduced to the maximum value of " _
           + CStr(maxPeriod) + ".", VbExclamation
    CurrentScript.AttValue("Period") = maxPeriod
  End If
End Sub

'==============================================================================================  
' End of script
'==============================================================================================  
