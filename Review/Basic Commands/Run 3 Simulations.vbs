'==========================================================================
' VBS-Script for Vissim 2022+
' Copyright (C) PTV AG, Sven Beller
' All rights reserved.
'==========================================================================

Option Explicit 

' Run 3 short simulation runs with different seeds to get some evaluations

' Please note: Controlling the simulation is not possible with integrated
'              (event-based) scripting, but only from outside Vissim.


' Ensure that all previous results are deleted in order to start clean
Dim simRun
for each simRun in Vissim.Net.SimulationRuns
Vissim.Net.SimulationRuns.RemoveSimulationRun(simRun)
next

' To maximize simulation speed ...
' ... activate QuickMode:
Vissim.Graphics.CurrentNetworkWindow.AttValue("QuickMode") = 1

' ... stop updating the Vissim workspace (network editors, lists, charts and other Vissim windows) 
'     while COM script controls the simulation:
Vissim.SuspendUpdateGUI

' Set the simulation time bounds
Vissim.Simulation.AttValue("SimPeriod") = 600	'[Simulation second]
Vissim.Simulation.AttValue("SimBreakAt") = 0  '[Simulation second]. If set to 0 then the simulation is not paused.

' Run the simulation three times
Dim simRunIndex
For simRunIndex = 1 To 3
	Vissim.Simulation.AttValue("RandSeed") = simRunIndex		' ensure that each simulation run is done with a different random seed
	Vissim.Simulation.RunContinuous
Next

' restore the Vissim workspace back to normal operation and deactivate quick mode
Vissim.ResumeUpdateGUI (true)
Vissim.Graphics.CurrentNetworkWindow.AttValue("QuickMode") = 0
