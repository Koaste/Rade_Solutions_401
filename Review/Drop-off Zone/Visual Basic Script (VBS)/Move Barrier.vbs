'==========================================================================
' VB Script 
' "Move Barrier"
' for use as integrated script with PTV Vissim example "Drop-off Zone"
'
' Copyright (c) Sven Beller, PTV AG.
' All rights reserved.
'==========================================================================

Option Explicit

' ---------------------------------------------------------------------------------------------
' Constants
' ---------------------------------------------------------------------------------------------
Const BARRIER_ALLOWANCE = 2       ' duration [s] between start of barrier opening until vehicle starts 

' ---------------------------------------------------------------------------------------------
' Declarations of global variables
' ---------------------------------------------------------------------------------------------
Dim simRes
Dim maxState					        ' max. state number for static 3D model of the barrier
Dim barrierLinkNo			        ' Number of the Vissim link/connector next to the barrier
Dim barrierObjNo			        ' Number of the static 3D model that represents the barrier  
Dim barrier						        ' The 3D model that represents the barrier
Dim barrierDwellTmDistr		    ' Number of the dwell time distribution associated with the stop sign at the barrier
Dim barrierDwellTmLowerBound  ' Lower bound of the dwell time distribution associated with the barrier stop sign
Dim currentScriptFileNoPath   ' name of the current script without the path information

' ---------------------------------------------------------------------------------------------
' Initialization
' ---------------------------------------------------------------------------------------------
Call Initialization   ' is called only once when the script is run for the first time and not thereafter


'==============================================================================================  
Sub Initialization
' General initialization, e.g. assigning values to the global variables
' Required globals: all globally declared variables

  simRes = vissim.simulation.AttValue("SimRes")

  ' Get the filename of the script
  dim pos
  pos = InStrRev(CurrentScriptFile, "\")
  currentScriptFileNoPath = Mid(CurrentScriptFile, pos+1)

  ' Get the script-associated UDAs
  barrierLinkNo = GetAndCheckScriptUDA ("RelLinkNo")
  if barrierLinkNo = 0 then exit sub
  
  barrierObjNo = GetAndCheckScriptUDA ("MainObjNo")
  if barrierObjNo = 0 then exit sub
  
  ' Associate the barrier object
  Set barrier = vissim.Net.Static3DModels.ItemByKey(barrierObjNo)
  maxState = barrier.AttValue("NumStates") - 1    ' first state is 0 (zero-based)
  barrier.AttValue("State") = 0		' start simulation with barrier closed

  ' Get the dwell time distribution of the stop sign 
  Dim stopSign
  Dim laneLinkNo
  for each stopSign in vissim.net.StopSigns
    if stopSign.AttValue("Lane\Link\No") = barrierLinkNo then
      barrierDwellTmDistr = stopSign.AttValue("DwellTmDistr(10)")
    end if
  next

  if barrierDwellTmDistr <= 0 then 
    msgbox "Dwell time distribution associated with the stop sign of the barrier was not found. Default distribution no. 1 chosen.", vbExclamation
    barrierDwellTmDistr = 1
  end if

  'Get lower bound of dwell time
  with vissim.Net.TimeDistributions.ItemByKey(barrierDwellTmDistr)
    if .AttValue("Type") = "EMPIRICAL" then
      barrierDwellTmLowerBound = .AttValue("LowerBound")
    else
      barrierDwellTmLowerBound = .AttValue("Mean") - 3 * .AttValue("StdDev")  ' only an approximation...
    end if
    
    if barrierDwellTmLowerBound < (maxState / simRes) then    ' if the dwell time is less than the min. required for this sim resolution...
      msgbox "The lower bound of dwell time distribution " + CStr(barrierDwellTmDistr) + " (which is associated with the stop signs) is too low for the barrier to fully open with the current setting of simulation resolution. Some vehicles may visually pass a closed barrier."
    end if
  end with

end Sub
  
'==============================================================================================  
'==============================================================================================  
'==============================================================================================  
Sub Main()
' Main program to be executed during the simulation
	
	Call BoundsCheck()
	Call MoveBarrier()	

End Sub

' ==============================================================================================  
Sub MoveBarrier()  
' Controls the barrier opening and closing process.
' Required globals: barrierLinkNo, barrier, maxState, BARRIER_ALLOWANCE
'
' The barrier starts opening as soon as the vehicle in front of the barrier has a remaining 
' dwell time of <BARRIER_ALLOWANCE> seconds or less.
' The barrier closes as soon as the vehicle has left the barrier link.
' It is assumed that never more than one vehicle is located on the barrier link.

  ' open barrier
  Dim veh
  For Each veh In vissim.Net.Links.ItemByKey(barrierLinkNo).Vehs
      ' open the barrier as soon as the dwell time of the vehicle is BARRIER_ALLOWANCE sec or less
      If veh.AttValue("DwellTm") > 0 And veh.AttValue("DwellTm") <= BARRIER_ALLOWANCE And barrier.AttValue("State") < maxState Then
          barrier.AttValue("State") = barrier.AttValue("State") + 1
      End If
  Next
  
  ' close barrier
  If vissim.Net.Links.ItemByKey(barrierLinkNo).Vehs.Count = 0 And barrier.AttValue("State") > 0 Then
      barrier.AttValue("State") = barrier.AttValue("State") - 1
  End If
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

' ==============================================================================================  
Sub BoundsCheck()
' Ensures that the script period is small enough for the script to run correctly.
' Required globals: barrierDwellTmLowerBound, BARRIER_ALLOWANCE, maxState, currentScriptFileNoPath
  
' As the script period may be changed during a simulation run, the bounds check is not only run 
' before sim start but every time the script is executed. 
' The script MoveBarrier needs to be executed at least as often as it needs to fully open the barrier 
' (<maxState> times (= state switches) within BARRIER_ALLOWANCE seconds)

  Dim maxPeriod
  if (barrierDwellTmLowerBound > 0) and (barrierDwellTmLowerBound < BARRIER_ALLOWANCE) then
    maxPeriod = barrierDwellTmLowerBound * simRes \ maxState
  else
    maxPeriod = BARRIER_ALLOWANCE * simRes \ maxState
  end if
  
  Dim scriptPeriod
  scriptPeriod = CurrentScript.AttValue("Period")

  if (maxPeriod = 0) then
    msgbox "You need to increase the simulation resolution in order for the script 'Move Barrier' to run correctly. The simulation will be stopped now.", vbCritical
    vissim.simulation.stop
  elseIf (scriptPeriod > maxPeriod) then 
    msgbox "The script period of " + CStr(scriptPeriod) + " for '" + currentScriptFileNoPath + "'" _ 
           + vbCrLf + "is too coarse for the currently defined" _
           + vbCrLf + "stopping time and barrier movement." _ 
           + vbCrLf + "Hence it is reduced to the maximum value of " _
           + CStr(maxPeriod) + ".", VbExclamation
    CurrentScript.AttValue("Period") = maxPeriod
  end if
End Sub

' ==============================================================================================  
' End of script
' ==============================================================================================  
  