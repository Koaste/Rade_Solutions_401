"""
#==========================================================================
# Python-Script for Vissim 9+
# Copyright (C) PTV AG. All rights reserved.
# Jochen Lohmiller 2016
# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
# Example for building a network through Vissim COM
#==========================================================================

# This script demonstrates how to add and remove network objects in Python.
# All 'Add' methods are descried in the PTV Vissim COM help, which can be accessed in the
# PTV Vissim menu 'Help - COM Help'.
# If you search for 'add*' in the COM Help, your find all 'Add' methods.
# For all objects where the key is the 'Number' attribute, the value of 0 means that Vissim
# will automatically choose a number that is not in use.

# There are also methods for removing objects. You can find the remove methods if you search
# for 'remove*' in the COM help.
# Example to remove a link: Vissim.Net.Links.RemoveLink(Vissim.Net.Links.ItemByKey(1))
"""

# COM-Server
import win32com.client as com

# Connecting the COM Server
Vissim = com.Dispatch("Vissim.Vissim")
Vissim.New()

PTVVissimInstallationPath = Vissim.AttValue("ExeFolder")    # directory of your PTV Vissim installation (where vissim.exe is located)

# For this example, units are set to Metric.
# Note: PTV Vissim coordinates are always in meters [m]
UnitCurrent = Vissim.Net.NetPara.AttValue('UnitLenShort')
UnitAttributes = ('UnitAccel', 'UnitLenLong', 'UnitLenShort', 'UnitLenVeryShort', 'UnitSpeed', 'UnitSpeedSmall')
for UnitAttrCurr in UnitAttributes:
    Vissim.Net.NetPara.SetAttValue(UnitAttrCurr, 0)         # 0: Metric [m], 1: Imperial [ft]

# Zoom Network Editor
Vissim.Graphics.CurrentNetworkWindow.ZoomTo(-300, -300, 300, 300)

#======================================
#======================================
# ADDING OBJECTS FOR VEHICULAR TRAFFIC
#======================================
#======================================

#--------------------------------------
# Base Data
#--------------------------------------
Vissim.Net.VehicleClasses.AddVehicleClass(0) # unsigned int Key
Vissim.Net.VehicleTypes.AddVehicleType(0) # unsigned int Key
Vissim.Net.DrivingBehaviors.AddDrivingBehavior(0) # unsigned int Key
Vissim.Net.LinkBehaviorTypes.AddLinkBehaviorType(0) # unsigned int Key
Vissim.Net.LinkBehaviorTypes.ItemByKey(1).VehClassDrivBehav.AddVehClassDrivingBehavior(Vissim.Net.VehicleClasses.ItemByKey(30)) # IVehicleClass* VehClass

Vissim.Net.DisplayTypes.AddDisplayType(0) # unsigned int Key
Vissim.Net.Levels.AddLevel(0) # unsigned int Key

#--------------------------------------
# Vehicle Compositions
#--------------------------------------
Vissim.Net.VehicleCompositions.AddVehicleComposition(0, []) # unsigned int Key, SAFEARRAY(VARIANT) VehCompRelFlows
Vissim.Net.VehicleCompositions.AddVehicleComposition(0, [Vissim.Net.VehicleTypes.ItemByKey(100), Vissim.Net.DesSpeedDistributions.ItemByKey(40)]) # unsigned int Key, SAFEARRAY(VARIANT) VehCompRelFlows
Vissim.Net.VehicleCompositions.AddVehicleComposition(9, [Vissim.Net.VehicleTypes.ItemByKey(100), Vissim.Net.DesSpeedDistributions.ItemByKey(40), Vissim.Net.VehicleTypes.ItemByKey(200), Vissim.Net.DesSpeedDistributions.ItemByKey(30)]) # unsigned int Key, SAFEARRAY(VARIANT) VehCompRelFlows
Vissim.Net.VehicleCompositions.ItemByKey(9).VehCompRelFlows.AddVehicleCompositionRelativeFlow(Vissim.Net.VehicleTypes.ItemByKey(300), Vissim.Net.DesSpeedDistributions.ItemByKey(25)) # IVehicleType* VehType, IDesSpeedDistribution* DesSpeedDistr

#--------------------------------------
# Functions
#--------------------------------------
Vissim.Net.DesAccelerationFunctions.AddDesAccelerationFunction(0, [0, 3, 2, 4,  120, 6, 5, 7,  160, 0, 0, 0]) # unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts [X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...]
Vissim.Net.DesDecelerationFunctions.AddDesDecelerationFunction(0, [0, -3, -4, -2,  120, -6, -7, -5,  160, 0, 0, 0]) # unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts [X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...]
Vissim.Net.MaxAccelerationFunctions.AddMaxAccelerationFunction(0, [0, 3, 2, 4,  120, 6, 5, 7,  160, 0, 0, 0]) # unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts [X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...]
Vissim.Net.MaxDecelerationFunctions.AddMaxDecelerationFunction(0, [0, -3, -4, -2,  120, -6, -7, -5,  160, 0, 0, 0]) # unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts [X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...]

#--------------------------------------
# Distributions
#--------------------------------------
Vissim.Net.DesSpeedDistributions.AddDesSpeedDistribution(0, [50, 0,  60, 0.1,  70, 0.9,  80, 1]) # unsigned int Key, SAFEARRAY(VARIANT) SpeedDistrDatPts [X1(speed), Y1(0...1), X2, Y2, ...]
Vissim.Net.PowerDistributions.AddPowerDistribution(0, [60, 0,  75, 0.5,  100, 1]) # unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts [X1(power), Y1(0...1), X2, Y2, ...]
Vissim.Net.WeightDistributions.AddWeightDistribution(0, [500, 0,  1000, 0.5,  2000, 1]) # unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts [X1(weight), Y1(0...1), X2, Y2, ...]
Vissim.Net.TimeDistributions.AddTimeDistributionNormal(0) # unsigned int Key
Vissim.Net.TimeDistributions.AddTimeDistributionEmpirical(0, [10, 0,  12, 0.75,  15, 1]) # unsigned int Key, SAFEARRAY(VARIANT) DurDistrDataPts [X1(duration), Y1(0...1), X2, Y2, ...]
Vissim.Net.LocationDistributions.AddLocationDistribution(0, [0, 0,  0.5, 0.3,  1, 1]) # unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts [X1(location), Y1(0...1), X2, Y2, ...]
Vissim.Net.DistanceDistributions.AddDistanceDistribution(0, [50, 0,  90, 0.2,  110, 0.8,  150, 1]) # unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts [X1(Distance), Y1(0...1), X2, Y2, ...]
Vissim.Net.OccupancyDistributions.AddOccupancyDistributionNormal(0) # unsigned int Key
Vissim.Net.OccupancyDistributions.AddOccupancyDistributionEmpirical(0, [1, 0,  2, 0.66,  3, 0.8,  5, 0.9,  9, 1]) # unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts [X1(Distance), Y1(0...1), X2, Y2, ...]
# Replace of all Distribution points (does only work with empirical distributions):
Vissim.Net.DesSpeedDistributions.ItemByKey(5).SpeedDistrDatPts.ReplaceAll([3.9, 0,  4.1, 1]) # SAFEARRAY(VARIANT) Elements

ColorDistribution = Vissim.Net.ColorDistributions.AddColorDistribution(0, ["44556677","ff000000"]) # unsigned int Key, SAFEARRAY(VARIANT) ColorDistrEl [Color1(StringHEX), Color2, ...]
ColorDistribution.ColorDistrEl.AddColorDistributionElement(0,"ff994400") # unsigned int Key, BSTR Color (StringHEX)
ColorDistribution.ColorDistrEl.ReplaceAll(["11223344", "ffffffff"]) # SAFEARRAY(VARIANT) ColorDistrEl [Color1(StringHEX), Color2, ...]

#--------------------------------------
# 2D/3D Models
#--------------------------------------
# Model2D3D.Model2D3DSegs.AddModel2D3DSegment(0, 'C:\\1.v3d') # unsigned int Key, BSTR File3D
ModelFile1 = PTVVissimInstallationPath + '\\3DModels\\Vehicles\\Road\\Car - Volkswagen Golf (2007).v3d'
ModelFile2 = PTVVissimInstallationPath + '\\3DModels\\Vehicles\\Road\\HGV - EU 04 Tractor.v3d'
ModelFile2a = PTVVissimInstallationPath + '\\3DModels\\Vehicles\\Road\\HGV - EU 04a Trailer.v3d'
ModelFile2b = PTVVissimInstallationPath + '\\3DModels\\Vehicles\\Road\\HGV - EU 04b Trailer.v3d'
Vissim.Net.Models2D3D.AddModel2D3D(0, [ModelFile1]) # unsigned int Key, SAFEARRAY(VARIANT) Model2D3DSegs [ModelFile1(String), ModelFile2, ...]
Model2D3D = Vissim.Net.Models2D3D.AddModel2D3D(0, [ModelFile2, ModelFile2a]) # unsigned int Key, SAFEARRAY(VARIANT) Model2D3DSegs [ModelFile1(String), ModelFile2, ...]
Model2D3D.Model2D3DSegs.AddModel2D3DSegment(0, ModelFile2b) # unsigned int Key, BSTR File3D
# Model2D3D.Model2D3DSegs.RemoveModel2D3DSegment(Model2D3D.Model2D3DSegs.ItemByKey(1))

Model2D3DDistribution = Vissim.Net.Model2D3DDistributions.AddModel2D3DDistribution(0, [Vissim.Net.Models2D3D.ItemByKey(1), Vissim.Net.Models2D3D.ItemByKey(2)]) # unsigned int Key, SAFEARRAY(VARIANT) Model2D3DDistrEl [Model1(IModel2D3D), Model2, ...]
Model2D3DDistribution.Model2D3DDistrEl.AddModel2D3DDistributionElement(0, Vissim.Net.Models2D3D.ItemByKey(3)) # unsigned int Key, IModel2D3D* Model2D3D
Model2D3DDistribution.Model2D3DDistrEl.ReplaceAll([Vissim.Net.Models2D3D.ItemByKey(5), Vissim.Net.Models2D3D.ItemByKey(6)]) # SAFEARRAY(VARIANT) Model2D3DDistrEl [Model1(IModel2D3D), Model2, ...]

#--------------------------------------
# TimeIntervalSet
#--------------------------------------
# Possible values for the 'TimeIntervalSet' enumeration along with their integer representation:
#   AreaBehaviorType       8
#   ManagedLanesFacility   9
#   PedestrianInput        5
#   PedestrianRoutePartial 6
#   PedestrianRouteStatic 11
#   VehicleInput           1
#   VehicleRouteParking    4
#   VehicleRoutePartial    3
#   VehicleRoutePartialPT  7
#   VehicleRouteStatic     2
TimeIntervalSet = 1   # 1 = VehicleInput
Vissim.Net.TimeIntervalSets.ItemByKey(TimeIntervalSet).TimeInts.AddTimeInterval(0) # unsigned int Key

#--------------------------------------
# Links
#--------------------------------------
# Input parameters of AddLink:
# 1: unsigned int Key   = attribute Number (No)         | example: 0 or 123
# 2: BSTR WktLinestring = attribution Points3D          | example: 'LINESTRING(PosX1 PosY1 PosZ1, PosX2 PosY2 PosZ2, ..., PosXn PosYn PosZn)' with Pos as double : PosZ is optional
# 3: LaneWidths         = number and widths of lanes    | example: [WidthLane1, WiidthLane2, ... WidthLaneN] as double
LinksEnter = [
    # unsigned int Key, BSTR WktLinestring 'LINESTRING(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)', SAFEARRAY(double) LaneWidths [WidthLane1, WiidthLane2, ... WidthLaneN]
    Vissim.Net.Links.AddLink(0, 'LINESTRING(-300 -2, -25 -2)', [3.5]),
    Vissim.Net.Links.AddLink(0, 'LINESTRING(2 -300, 2 -25)', [3.5]),
    Vissim.Net.Links.AddLink(0, 'LINESTRING(300 2, 25 2)', [3.5]),
    Vissim.Net.Links.AddLink(0, 'LINESTRING(-2 300, -2 25)', [3.5]),
]

# Opposite direction:  #ILink* Link, unsigned int NumberOfLanes
LinksExit = [Vissim.Net.Links.GenerateOppositeDirection(LinksEnter[cnt_Enter], 1) for cnt_Enter in range(len(LinksEnter))]

#--------------------------------------
# Connectors
#--------------------------------------
# Input parameters of AddConnector:
# 1: unsigned int Key   = attribute Number (No)         | example: 0 or 123
# 2: ILane* FromLane                                    | example: Vissim.Net.Links.ItemByKey(1).Lanes.ItemByKey(1)
# 3: double FromPos                                     | example: 200
# 4: ILane* ToLane                                      | example: Vissim.Net.Links.ItemByKey(2).Lanes.ItemByKey(1)
# 5: double ToPos                                       | example: 0
# 6: unsigned int NumberOfLanes                         | example: 1, 2, ...
# 7: BSTR WktLinestring = attribute Points3D            | example: 'LINESTRING(PosX2 PosY2 PosZ2, ..., PosXn-1 PosYn-1 PosZn-1)' with Pos as double; PosZ is optional; Pos1 and Posn automatically created; No additional points: 'LINESTRING EMPTY'
for cnt_Enter in range(len(LinksEnter)):
    for cnt_Exit in range(len(LinksExit)):
        if cnt_Enter != cnt_Exit:
            Vissim.Net.Links.AddConnector(0, LinksEnter[cnt_Enter].Lanes.ItemByKey(1), 275, LinksExit[cnt_Exit].Lanes.ItemByKey(1), 0, 1, 'LINESTRING EMPTY') # unsigned int Key, ILane* FromLane, double FromPos, ILane* ToLane, double ToPos, unsigned int NumberOfLanes, BSTR WktLinestring

# alternative Route:
LinksAlternative = Vissim.Net.Links.AddLink(0, 'LINESTRING(-150.000 -9.000 0.000,-113.009 -15.641 0.000,-75.582 -20.384 0.000,-37.864 -23.231 0.000,-0.000 -24.179 0.000,37.864 -23.231 0.000,75.582 -20.384 0.000,113.009 -15.641 0.000,150.000 -9.000 0.000)', [3.5]) # unsigned int Key, BSTR WktLinestring, SAFEARRAY(double) LaneWidths
ConnToAlternative = Vissim.Net.Links.AddConnector(0, LinksEnter[0].Lanes.ItemByKey(1), 130, LinksAlternative.Lanes.ItemByKey(1), 0, 1, 'LINESTRING(-166.600 -4.181 0.000,-163.323 -5.284 0.000,-160.114 -6.775 0.000,-156.913 -8.368 0.000,-153.664 -9.779 0.000)') # unsigned int Key, ILane* FromLane, double FromPos, ILane* ToLane, double ToPos, unsigned int NumberOfLanes, BSTR WktLinestring
ConnFromAlternative = Vissim.Net.Links.AddConnector(0, LinksAlternative.Lanes.ItemByKey(1), LinksAlternative.AttValue('Length2D'), LinksExit[2].Lanes.ItemByKey(1), 145, 1, 'LINESTRING(153.741 -9.595 0.000,156.974 -7.762 0.000,160.112 -5.631 0.000,163.259 -3.606 0.000,166.520 -2.094 0.000)') # unsigned int Key, ILane* FromLane, double FromPos, ILane* ToLane, double ToPos, unsigned int NumberOfLanes, BSTR WktLinestring
LinksAlternative.Lanes.AddLane(0, 3.5) # unsigned int Key, double Width

#--------------------------------------
# Vehicle Inputs
#--------------------------------------
for cnt_Enter in range(len(LinksEnter)):
    VehInput = Vissim.Net.VehicleInputs.AddVehicleInput(0, LinksEnter[cnt_Enter]) # unsigned int Key, Link,
    VehInput.SetAttValue('Volume(1)', 500)

#--------------------------------------
# Parking Lot
#--------------------------------------
ParkingLot1 = Vissim.Net.ParkingLots.AddParkingLot(1, LinksExit[2].Lanes.ItemByKey(1), 186) # unsigned int Key, ILane* Lane, double Pos

Vissim.Net.Zones.AddZone(0) # unsigned int Key
Vissim.Net.ParkingLots.AddAbstractParkingLot(0, LinksExit[3], 186, Vissim.Net.Zones.ItemByKey(1)) # unsigned int Key, ILink* Link, double Pos, IZone* Zone

#--------------------------------------
# Matrices
#--------------------------------------
Vissim.Net.Matrices.AddMatrix(0) # unsigned int Key

#--------------------------------------
# Vehicle Routing Decisions
#--------------------------------------
VehRoutDesSta = Vissim.Net.VehicleRoutingDecisionsStatic.AddVehicleRoutingDecisionStatic(0, LinksEnter[0], 100) # unsigned int Key, ILink* Link,, double Pos
VehRoutDesPar = Vissim.Net.VehicleRoutingDecisionsPartial.AddVehicleRoutingDecisionPartial(0, LinksEnter[0], 102, LinksExit[2], 202) # unsigned int Key, ILink* Link, double Pos, ILink* DestLink, double DestPos
VehRoutDesPPT = Vissim.Net.VehicleRoutingDecisionsPartialPT.AddVehicleRoutingDecisionPartialPT(0, LinksEnter[0], 104, LinksExit[2], 204) # unsigned int Key, ILink* Link, double Pos, ILink* DestLink, double DestPos
VehRoutDesPark = Vissim.Net.VehicleRoutingDecisionsParking.AddVehicleRoutingDecisionParking(0, LinksEnter[0], 106) # unsigned int Key, ILink* Link, double Pos
VehRoutDesML = Vissim.Net.VehicleRoutingDecisionsManagedLanes.AddVehicleRoutingDecisionManagedLanes(0, LinksEnter[0], 108, LinksExit[2], 208) # unsigned int Key, ILink* Link, double Pos, ILink* DestLink, double DestPos
VehRoutDesDyn = Vissim.Net.VehicleRoutingDecisionsDynamic.AddVehicleRoutingDecisionDynamic(0, LinksEnter[0], 110) # unsigned int Key, ILink* Link, double Pos
VehRoutDesClose = Vissim.Net.VehicleRoutingDecisionsClosure.AddVehicleRoutingDecisionClosure(0, LinksEnter[0], 112) # unsigned int Key, ILink* Link, double Pos

#--------------------------------------
# Vehicle Routes
#--------------------------------------
VehRoutSta1 = VehRoutDesSta.VehRoutSta.AddVehicleRouteStatic(0, LinksExit[2], 200) # unsigned int Key, ILink* DestLink, double DestPos
VehRoutSta2 = VehRoutDesSta.VehRoutSta.AddVehicleRouteStatic(0, LinksExit[2], 200) # unsigned int Key, ILink* DestLink, double DestPos
VehRoutSta2.UpdateLinkSequence([LinksAlternative, 20]) # SAFEARRAY(VARIANT) RouteSegments Recomputes the link sequence. RouteSegments must be a list of each link and its offset that should be part of the route.
VehRoutSta2.UpdateLinkSequence([ConnToAlternative, 1, LinksAlternative, 20, ConnFromAlternative, 1]) # SAFEARRAY(VARIANT) RouteSegments

VehRoutPar = VehRoutDesPar.VehRoutPart.AddVehicleRoutePartial(0) # unsigned int Key
VehRoutDesPPT.VehRoutPartPT.AddVehicleRoutePartialPT(0) # unsigned int Key
VehRoutDesPark.VehRoutPark.AddVehicleRouteParking(0, ParkingLot1) # unsigned int Key, IParkingLot* ParkLot
VehRoutDesML.VehRoutMngLns.AddVehicleRouteManagedLanes(0, 1) # unsigned int Key, enum ManagedLaneType Type [1: Managed; 2: General purpose]
VehRoutDesClose.VehRoutClos.AddVehicleRouteClosure(0, LinksExit[2], 204) # unsigned int Key, ILink* DestLink, double DestPos

Lane1 = LinksEnter[0].Lanes.ItemByKey(1)
Lane2 = LinksEnter[1].Lanes.ItemByKey(1)
Lane3 = LinksEnter[2].Lanes.ItemByKey(1)
Lane4 = LinksEnter[3].Lanes.ItemByKey(1)

#--------------------------------------
# Stop Sign
#--------------------------------------
Vissim.Net.StopSigns.AddStopSign(0, Lane1, 274) # unsigned int Key, ILane* Lane, double Pos

#--------------------------------------
# Priority Rule
#--------------------------------------
PriRule = Vissim.Net.PriorityRules.AddPriorityRule(0, Lane1, 270, [Lane2, 270, Lane3, 270]) # unsigned int Key, ILane* Lane, double Pos, SAFEARRAY(VARIANT) ConflictMarkers (ConflictMarkers must be a list that specifies for each ConflictMarker alternating its lane and position)
PriRule.ConflictMarkers.AddConflictMarker(0, Lane4, 270) #unsigned int Key, ILane* Lane, double Pos

#--------------------------------------
# Desired Speed Decision
#--------------------------------------
Vissim.Net.DesSpeedDecisions.AddDesSpeedDecision(0, Lane1, 10) # int Key, ILane* Lane, double Pos

#--------------------------------------
# Reduced Speed Area
#--------------------------------------
Vissim.Net.ReducedSpeedAreas.AddReducedSpeedArea(0, Lane1, 10) # unsigned int Key, ILane* Lane, double Pos

#--------------------------------------
# Node
#--------------------------------------
Vissim.Net.Nodes.AddNode(0, 'POLYGON((-30 -30, 30 -30, 30 30, -30 30, -30 -30))') # unsigned int Key, BSTR WktPolygon

#--------------------------------------
# Public Transport
#--------------------------------------
Vissim.Net.PTStops.AddPTStop(0, Lane1, 150) # unsigned int Key, ILane* Lane, double Pos
PTLine = Vissim.Net.PTLines.AddPTLine(0, LinksEnter[0], LinksExit[2], 200) # unsigned int Key, ILink* EntryLink, ILink* DestLink, double DestPos
PTLine.DepTimes.AddDepartureTime(0, 123) # unsigned int Key, double Dep
PTLine.DepTimes.AddDepartureTime(0, 246) # unsigned int Key, double Dep
PTLine.UpdateLinkSequence([LinksAlternative, 20]) # SAFEARRAY(VARIANT) RouteSegments Recomputes the link sequence. RouteSegments must be a list of each link and its offset that should be part of the route.
# PTLine.UpdateLinkSequence([ConnToAlternative, 1, LinksAlternative, 20, ConnFromAlternative, 1]) # SAFEARRAY(VARIANT)

#--------------------------------------
# Static 3D Models
#--------------------------------------
Static3DModelFile = PTVVissimInstallationPath + '\\3DModels\\Vegetation\\Tree 04.v3d'
Vissim.Net.Static3DModels.AddStatic3DModel(0, Static3DModelFile, 'Point(-100, -100, 0)') # unsigned int Key, BSTR ModelFilename, BSTR PosWktPoint3D

#--------------------------------------
# Background Image
#--------------------------------------
BackgroundImageFile = PTVVissimInstallationPath + '\\3DModels\\Textures\\Marsh.png'
Vissim.Net.BackgroundImages.AddBackgroundImage(0, BackgroundImageFile, 'Point(100, 100)', 'Point(-100, -100)') # unsigned int Key, BSTR PathFilename, BSTR PosTRWktPoint, BSTR PosBLWktPoint

#--------------------------------------
# Presentation
#--------------------------------------
CamPos = Vissim.Net.CameraPositions.AddCameraPosition(0, 'Point(-100, -100, 100)') # unsigned int Key, BSTR WktPoint3D
CamPos.SetAttValue('Name', 'Position 1 (added by COM)')
Storyboard = Vissim.Net.Storyboards.AddStoryboard(0) # unsigned int Key
Storyboard.Keyframes.AddKeyframe(0) # unsigned int Key

#--------------------------------------
# Section
#--------------------------------------
Vissim.Net.Sections.AddSection(0, 'POLYGON((-200 -150,  200 -150,  200 150,  -200 150))') # unsigned int Key, BSTR WktPolygon

#--------------------------------------
# Signal Controller
#--------------------------------------
SignalController = Vissim.Net.SignalControllers.AddSignalController(0) # unsigned int Key
SignalController.SGs.AddSignalGroup(0) # unsigned int Key

#--------------------------------------
# Detector
#--------------------------------------
Vissim.Net.Detectors.AddDetector(0, Lane1, 250) # unsigned int Key, Lane, double Pos

#--------------------------------------
# Signal Head
#--------------------------------------
Vissim.Net.SignalHeads.AddSignalHead(0, Lane1, 274) # unsigned int Key, ILane* Lane, double Pos

#--------------------------------------
# 3D Traffic Signal
#--------------------------------------
TrafficSignal3D = Vissim.Net.Signals3D.AddTrafficSignal3D(0, 'POINT(0 0)') # unsigned int Key, WktPos3D As String
TrafficSignal3D.SigHeads.AddSignalHead3D(0) # unsigned int Key
TrafficSignal3D.SigArms.AddSignalArm3D(0) # unsigned int Key
TrafficSignal3D.TrafficSigns.AddTrafficSign3D(0) # unsigned int Key
TrafficSignal3D.Streetlights.AddStreetlight3D(0) # unsigned int Key

#--------------------------------------
# Pavement Marking
#--------------------------------------
Vissim.Net.PavementMarkings.AddPavementMarking(0, Lane1, 215) # unsigned int Key, ILane* Lane, double Pos

#--------------------------------------
# Vehicles Evaluations
#--------------------------------------
Vissim.Net.DataCollectionPoints.AddDataCollectionPoint(0, Lane1, 230) # unsigned int Key, ILane* Lane, double Pos
Vissim.Net.VehicleTravelTimeMeasurements.AddVehicleTravelTimeMeasurement(0, LinksEnter[1], 50, LinksExit[0], 200) # unsigned int Key, ILink* StartLink, double StartPos, ILink* EndLink, double EndPos
Vissim.Net.QueueCounters.AddQueueCounter(0, LinksEnter[0], 123) # unsigned int Key, ILink* Link, double Pos
Vissim.Net.DelayMeasurements.AddDelayMeasurement(0) # unsigned int Key




#=============================================
#=============================================
# ADDING OBJECTS SPECIFICALLY FOR PEDESTRIANS
#=============================================
#=============================================

#--------------------------------------
# Pedestrian Base Data
#--------------------------------------
Vissim.Net.PedestrianTypes.AddPedestrianType(0) # unsigned int Key
Vissim.Net.PedestrianClasses.AddPedestrianClass(0) # unsigned int Key
Vissim.Net.WalkingBehaviors.AddWalkingBehavior(0) # unsigned int Key
AreaBehaviorType = Vissim.Net.AreaBehaviorTypes.AddAreaBehaviorType(0) # unsigned int Key
AreaBehaviorType.AreaBehavTypeElements.AddAreaBehaviorTypeElement(Vissim.Net.PedestrianClasses.ItemByKey(30)) # IPedestrianClass* PedClass

#--------------------------------------
# Pedestrian Compositions
#--------------------------------------
Vissim.Net.PedestrianCompositions.AddPedestrianComposition(0, []) # unsigned int Key, SAFEARRAY(VARIANT) PedCompRelFlows
Vissim.Net.PedestrianCompositions.AddPedestrianComposition(9, [Vissim.Net.PedestrianTypes.ItemByKey(100), Vissim.Net.DesSpeedDistributions.ItemByKey(5), Vissim.Net.PedestrianTypes.ItemByKey(200), Vissim.Net.DesSpeedDistributions.ItemByKey(5)]) # unsigned int Key, SAFEARRAY(VARIANT) PedCompRelFlows
Vissim.Net.PedestrianCompositions.ItemByKey(9).PedCompRelFlows.AddPedestrianCompositionRelativeFlow(Vissim.Net.PedestrianTypes.ItemByKey(300), Vissim.Net.DesSpeedDistributions.ItemByKey(5)) # IPedestrianType* PedType, IDesSpeedDistribution* DesSpeedDistr

#--------------------------------------
# Areas
#--------------------------------------
# Input parameters of AddArea:
# 1: unsigned int Key   = attribute Number (No)          | example: 0 or 123
# 2: BSTR WktPolygon    = attribute Points               | example: 'POLYGON(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)' with Pos as double

# Example: Area1 = Vissim.Net.Areas.AddArea(0, 'POLYGON((0 50, 50 50, 50 100, 0 100))') # unsigned int Key, BSTR WktPolygon

# Create some areas:
xydim = (10, 10)
StartPosAll = ((50, 50), (50, 60), (50, 70), (50, 80), (50, 90), (60, 60), (70, 60), (70, 70), (70, 80), (60, 80))
Area = []
for cnt, StartPos in enumerate(StartPosAll):
    x1 = str(StartPos[0])
    x2 = str(StartPos[0] + xydim[0])
    y1 = str(StartPos[1])
    y2 = str(StartPos[1] + xydim[1])
    Polyg = 'POLYGON((' + x1 + ' ' + y1 +',' + x1 + ' ' + y2 + ',' + x2 + ' ' + y2 + ',' + x2 + ' ' + y1 + '))'
    Area.append(Vissim.Net.Areas.AddArea(0, Polyg)) # unsigned int Key, BSTR WktPolygon

#--------------------------------------
# Obstacle & Ramp
#--------------------------------------
Vissim.Net.Obstacles.AddObstacle(0, 'POLYGON((55 72.5, 65 72.5, 65 77.5, 55 77.5))') # unsigned int Key, BSTR WktPolygon
Ramp = Vissim.Net.Ramps.AddRamp(0, 'POLYGON((50 100,  50 110,  60 110,  60 100))') # unsigned int Key, BSTR WktPolygon
Area.append(Vissim.Net.Areas.AddArea(0, 'POLYGON((50 110,  60 110,  60 120,  50 120))')) # unsigned int Key, BSTR WktPolygon

#--------------------------------------
# Pedestrian Input
#--------------------------------------
Vissim.Net.PedestrianInputs.AddPedestrianInput(0, Area[0]) # unsigned int Key, IArea* Area

#--------------------------------------
# Pedestrian Routing Decisions & Routes
#--------------------------------------
PedRouteDesSta1 = Vissim.Net.PedestrianRoutingDecisionsStatic.AddPedestrianRoutingDecisionStatic(0, Area[0]) # unsigned int Key, IArea* Area
PedRouteSta1 = PedRouteDesSta1.PedRoutSta.AddPedestrianRouteStaticOnArea(0, Area[1]) # unsigned int Key, IArea* Area
PedRouteSta1.PedRoutLoc.AddPedestrianRouteLocationOnArea(0, Area[2]) # unsigned int Key, IArea* Area | => will become next route end
PedRouteSta1.PedRoutLoc.AddPedestrianRouteLocationOnArea(0, Area[5]) # unsigned int Key, IArea* Area | => will become next route end
PedRouteSta1.PedRoutLoc.AddPedestrianRouteLocationOnArea(0, Area[3]) # unsigned int Key, IArea* Area | => will become next route end
PedRouteSta1.PedRoutLoc.AddPedestrianRouteLocationOnArea(0, Area[4]) # unsigned int Key, IArea* Area | => will become next route end

PedRouteSta2 = PedRouteDesSta1.PedRoutSta.AddPedestrianRouteStaticOnRamp(0, Ramp)  # unsigned int Key, IRamp* PedRoutLoc
PedRouteSta3 = PedRouteDesSta1.PedRoutSta.AddPedestrianRouteStaticOnArea(0, Area[9]) # unsigned int Key, IArea* Area
PedRouteSta3.PedRoutLoc.AddPedestrianRouteLocationOnRamp(0, Ramp) # unsigned int Key, IRamp* PedRoutLoc | => will become next route end

PedRouteDesPart1 = Vissim.Net.PedestrianRoutingDecisionsPartial.AddPedestrianRoutingDecisionPartial(0, Area[1]) # unsigned int Key, IArea* Area
PedRoutePart1 = PedRouteDesPart1.PedRoutPart.AddPedestrianRoutePartialOnArea(0, Area[7]) # unsigned int Key, IArea* Area
PedRoutePart1.PedRoutLoc.AddPedestrianRouteLocationOnArea(0, Area[3]) # unsigned int Key, IArea* Area | => will become next route end
PedRoutePart2 = PedRouteDesPart1.PedRoutPart.AddPedestrianRoutePartialOnArea(0, Area[3]) # unsigned int Key, IArea* Area

PedRouteDesPart2 = Vissim.Net.PedestrianRoutingDecisionsPartial.AddPedestrianRoutingDecisionPartial(0, Area[1]) # unsigned int Key, IArea* Area
PedRoutePart2 = PedRouteDesPart2.PedRoutPart.AddPedestrianRoutePartialOnRamp(0, Ramp)  # unsigned int Key, IRamp* PedRoutLoc

#--------------------------------------
# Pedestrian Evaluation
#--------------------------------------
Vissim.Net.PedestrianTravelTimeMeasurements.AddPedestrianTravelTimeMeasurement(0, Area[0], Area[4]) # unsigned int Key, IArea* StartArea, IArea* EndArea

#======================================
# END OF SCRIPT
#======================================
