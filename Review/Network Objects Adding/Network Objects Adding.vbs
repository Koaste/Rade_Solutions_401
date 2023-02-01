'==========================================================================
' VBS-Script for Vissim 9+
' Copyright (C) PTV AG. All rights reserved.
' Jochen Lohmiller 2016
' -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
' Example for building a network through Vissim COM
'==========================================================================

' This script demonstrates how to add and remove network objects in Python.
' All 'Add' methods are descried in the PTV Vissim COM help, which can be accessed in the PTV Vissim menu 'Help - COM Help'.
' If you search for 'add*' in the COM Help, your find all 'Add' methods.
' For all objects where the key is the 'Number' attribute, the value of 0 means that Vissim will automatically choose a number that is not in use.

' There are also methods for removing objects. You can find the remove methods if you search for 'remove*' in the COM help.
' Example to remove a link: Vissim.Net.Links.RemoveLink(Vissim.Net.Links.ItemByKey(1))

Option Explicit

Dim PTVVissimInstallationPath
PTVVissimInstallationPath = Vissim.AttValue("ExeFolder")    ' directory of your PTV Vissim installation (where vissim.exe is located)

' For this example, units are set to Metric:
' Note: PTV Vissim Coordinates are always in meters (m)
Dim UnitCurrent
Dim UnitAttributes
Dim i
UnitCurrent = Vissim.Net.NetPara.AttValue("UnitLenShort")
UnitAttributes = Array ("UnitAccel", "UnitLenLong", "UnitLenShort", "UnitLenVeryShort", "UnitSpeed", "UnitSpeedSmall")

for i = 0 to Ubound(UnitAttributes)
    Vissim.Net.NetPara.AttValue(UnitAttributes(i)) = 0 ' 0: Metric (m), 1: Imperial (ft)
Next

' Zoom Network Editor
Vissim.Graphics.CurrentNetworkWindow.ZoomTo -300, -300, 300, 300

'======================================
'======================================
' ADDING OBJECTS FOR VEHICULAR TRAFFIC
'======================================
'======================================

'--------------------------------------
' Base Data
'--------------------------------------

Vissim.Net.VehicleClasses.AddVehicleClass(0) ' unsigned int Key
Vissim.Net.VehicleTypes.AddVehicleType(0) ' unsigned int Key
Vissim.Net.DrivingBehaviors.AddDrivingBehavior(0) ' unsigned int Key
Vissim.Net.LinkBehaviorTypes.AddLinkBehaviorType(0) ' unsigned int Key
Vissim.Net.LinkBehaviorTypes.ItemByKey(1).VehClassDrivBehav.AddVehClassDrivingBehavior Vissim.Net.VehicleClasses.ItemByKey(30) ' IVehicleClass* VehClass

Vissim.Net.DisplayTypes.AddDisplayType(0) ' unsigned int Key
Vissim.Net.Levels.AddLevel(0) ' unsigned int Key

'--------------------------------------
' Vehicle Compositions
'--------------------------------------
Dim VehCompRelFlows1(1), VehCompRelFlows2(3), VehicleComposition
Set VehCompRelFlows1(0) = Vissim.Net.VehicleTypes.ItemByKey(100)
Set VehCompRelFlows1(1) = Vissim.Net.DesSpeedDistributions.ItemByKey(40)
Set VehCompRelFlows2(0) = Vissim.Net.VehicleTypes.ItemByKey(100)
Set VehCompRelFlows2(1) = Vissim.Net.DesSpeedDistributions.ItemByKey(40)
Set VehCompRelFlows2(2) = Vissim.Net.VehicleTypes.ItemByKey(200)
Set VehCompRelFlows2(3) = Vissim.Net.DesSpeedDistributions.ItemByKey(30)

Vissim.Net.VehicleCompositions.AddVehicleComposition 0, VehCompRelFlows1 ' unsigned int Key, SAFEARRAY(VARIANT) VehCompRelFlows
Set VehicleComposition = Vissim.Net.VehicleCompositions.AddVehicleComposition(0, VehCompRelFlows2) ' unsigned int Key, SAFEARRAY(VARIANT) VehCompRelFlows
VehicleComposition.VehCompRelFlows.AddVehicleCompositionRelativeFlow Vissim.Net.VehicleTypes.ItemByKey(300), Vissim.Net.DesSpeedDistributions.ItemByKey(25) ' IVehicleType* VehType, IDesSpeedDistribution* DesSpeedDistr

'--------------------------------------
' Functions
'--------------------------------------
Dim AccelFuncDataPts1(11), AccelFuncDataPts2(11)
AccelFuncDataPts1(0) = 0
AccelFuncDataPts1(1) = 3
AccelFuncDataPts1(2) = 2
AccelFuncDataPts1(3) = 4
AccelFuncDataPts1(4) = 120
AccelFuncDataPts1(5) = 6
AccelFuncDataPts1(6) = 5
AccelFuncDataPts1(7) = 7
AccelFuncDataPts1(8) = 160
AccelFuncDataPts1(9) = 0
AccelFuncDataPts1(10) = 0
AccelFuncDataPts1(11) = 0
AccelFuncDataPts2(0) = 0
AccelFuncDataPts2(1) = -3
AccelFuncDataPts2(2) = -4
AccelFuncDataPts2(3) = -2
AccelFuncDataPts2(4) = 120
AccelFuncDataPts2(5) = -6
AccelFuncDataPts2(6) = -7
AccelFuncDataPts2(7) = -5
AccelFuncDataPts2(8) = 160
AccelFuncDataPts2(9) = 0
AccelFuncDataPts2(10) = 0
AccelFuncDataPts2(11) = 0
Vissim.Net.DesAccelerationFunctions.AddDesAccelerationFunction 0, AccelFuncDataPts1 ' unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts (X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...)
Vissim.Net.DesDecelerationFunctions.AddDesDecelerationFunction 0, AccelFuncDataPts2 ' unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts (X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...)
Vissim.Net.MaxAccelerationFunctions.AddMaxAccelerationFunction 0, AccelFuncDataPts1 ' unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts (X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...)
Vissim.Net.MaxDecelerationFunctions.AddMaxDecelerationFunction 0, AccelFuncDataPts2 ' unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts (X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...)

'--------------------------------------
' Distributions
'--------------------------------------
Dim SpeedDistrDatPts1(7)
SpeedDistrDatPts1(0) = 50
SpeedDistrDatPts1(1) = 0
SpeedDistrDatPts1(2) = 60
SpeedDistrDatPts1(3) = 0.1
SpeedDistrDatPts1(4) = 70
SpeedDistrDatPts1(5) = 0.9
SpeedDistrDatPts1(6) = 80
SpeedDistrDatPts1(7) = 1
Vissim.Net.DesSpeedDistributions.AddDesSpeedDistribution 0, SpeedDistrDatPts1 ' unsigned int Key, SAFEARRAY(VARIANT) SpeedDistrDatPts (X1(speed), Y1(0...1), X2, Y2, ...)
Dim DistrDataPtsPower(5)
DistrDataPtsPower(0) = 500
DistrDataPtsPower(1) = 0
DistrDataPtsPower(2) = 1000
DistrDataPtsPower(3) = 0.5
DistrDataPtsPower(4) = 2000
DistrDataPtsPower(5) = 1
Vissim.Net.PowerDistributions.AddPowerDistribution 0, DistrDataPtsPower ' unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts (X1(power), Y1(0...1), X2, Y2, ...)
Dim DistrDataPtsWeigth(5)
DistrDataPtsWeigth(0) = 500
DistrDataPtsWeigth(1) = 0
DistrDataPtsWeigth(2) = 1000
DistrDataPtsWeigth(3) = 0.5
DistrDataPtsWeigth(4) = 2000
DistrDataPtsWeigth(5) = 1
Vissim.Net.WeightDistributions.AddWeightDistribution 0, DistrDataPtsWeigth ' unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts (X1(weight), Y1(0...1), X2, Y2, ...)
Dim DurDistrDataPts(5)
DurDistrDataPts(0) = 10
DurDistrDataPts(1) = 0
DurDistrDataPts(2) = 12
DurDistrDataPts(3) = 0.75
DurDistrDataPts(4) = 15
DurDistrDataPts(5) = 1
Vissim.Net.TimeDistributions.AddTimeDistributionEmpirical 0, DurDistrDataPts ' unsigned int Key, SAFEARRAY(VARIANT) DurDistrDataPts (X1(duration), Y1(0...1), X2, Y2, ...)
Vissim.Net.TimeDistributions.AddTimeDistributionNormal(0) ' unsigned int Key
Dim DistrDataPtsLocation(5)
DistrDataPtsLocation(0) = 0
DistrDataPtsLocation(1) = 0
DistrDataPtsLocation(2) = 0.5
DistrDataPtsLocation(3) = 0.3
DistrDataPtsLocation(4) = 1
DistrDataPtsLocation(5) = 1
Vissim.Net.LocationDistributions.AddLocationDistribution 0, DistrDataPtsLocation' unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts (X1(location), Y1(0...1), X2, Y2, ...)
Dim DistrDataPtsDistance(7)
DistrDataPtsDistance(0) = 50
DistrDataPtsDistance(1) = 0
DistrDataPtsDistance(2) = 90
DistrDataPtsDistance(3) = 0.2
DistrDataPtsDistance(4) = 110
DistrDataPtsDistance(5) = 0.8
DistrDataPtsDistance(6) = 150
DistrDataPtsDistance(7) = 1
Vissim.Net.DistanceDistributions.AddDistanceDistribution 0, DistrDataPtsDistance ' unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts (X1(Distance), Y1(0...1), X2, Y2, ...)
Dim DistrDataPtsOccupancy(9)
DistrDataPtsOccupancy(0) = 1
DistrDataPtsOccupancy(1) = 0
DistrDataPtsOccupancy(2) = 2
DistrDataPtsOccupancy(3) = 0.66
DistrDataPtsOccupancy(4) = 3
DistrDataPtsOccupancy(5) = 0.8
DistrDataPtsOccupancy(6) = 5
DistrDataPtsOccupancy(7) = 0.9
DistrDataPtsOccupancy(8) = 9
DistrDataPtsOccupancy(9) = 1
Vissim.Net.OccupancyDistributions.AddOccupancyDistributionEmpirical 0, DistrDataPtsOccupancy ' unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts (X1(Distance), Y1(0...1), X2, Y2, ...)
Vissim.Net.OccupancyDistributions.AddOccupancyDistributionNormal(0) ' unsigned int Key

' Replace all distribution points (does work only with empirical distributions):
Dim SpeedDistrDatPts2(3)
SpeedDistrDatPts2(0) = 3.9
SpeedDistrDatPts2(1) = 0
SpeedDistrDatPts2(2) = 4.1
SpeedDistrDatPts2(3) = 1
Vissim.Net.DesSpeedDistributions.ItemByKey(5).SpeedDistrDatPts.ReplaceAll SpeedDistrDatPts2 ' SAFEARRAY(VARIANT) Elements

Dim ColorDistribution, ColorDistrEl1(1)
ColorDistrEl1(0) = "44556677"
ColorDistrEl1(1) = "ff000000"
Set ColorDistribution = Vissim.Net.ColorDistributions.AddColorDistribution(0, ColorDistrEl1) ' unsigned int Key, SAFEARRAY(VARIANT) ColorDistrEl (Color1(StringHEX), Color2, ...)
ColorDistribution.ColorDistrEl.AddColorDistributionElement 0, "ff994400" ' unsigned int Key, BSTR Color (StringHEX)
Dim ColorDistrEl2(1)
ColorDistrEl2(0) = "11223344"
ColorDistrEl2(1) = "ffffffff"
ColorDistribution.ColorDistrEl.ReplaceAll ColorDistrEl2 ' SAFEARRAY(VARIANT) ColorDistrEl (Color1(StringHEX), Color2, ...)

Dim Model2D3DDistribution, Model2D3DDistrEl1(1)
Set Model2D3DDistrEl1(0) = Vissim.Net.Models2D3D.ItemByKey(1)
Set Model2D3DDistrEl1(1) = Vissim.Net.Models2D3D.ItemByKey(2)
Dim Model2D3DDistrEl2(1)
Set Model2D3DDistrEl2(0) = Vissim.Net.Models2D3D.ItemByKey(5)
Set Model2D3DDistrEl2(1) = Vissim.Net.Models2D3D.ItemByKey(6)
Set Model2D3DDistribution = Vissim.Net.Model2D3DDistributions.AddModel2D3DDistribution(0, Model2D3DDistrEl1) ' unsigned int Key, SAFEARRAY(VARIANT) Model2D3DDistrEl (Model1(IModel2D3D), Model2, ...)
Model2D3DDistribution.Model2D3DDistrEl.AddModel2D3DDistributionElement 0, Vissim.Net.Models2D3D.ItemByKey(3) ' unsigned int Key, IModel2D3D* Model2D3D
Model2D3DDistribution.Model2D3DDistrEl.ReplaceAll(Model2D3DDistrEl2) ' SAFEARRAY(VARIANT) Model2D3DDistrEl (Model1(IModel2D3D), Model2, ...)

'--------------------------------------
' 2D/3D Models
'--------------------------------------
Dim ModelFile1(0), ModelFile2(0), ModelFile2a, ModelFile2b, Model2D3DSegs(1), Model2D3D
' Model2D3D.Model2D3DSegs.AddModel2D3DSegment(0, "C:\\1.v3d") ' unsigned int Key, BSTR File3D
ModelFile1(0) = PTVVissimInstallationPath + "\\3DModels\\Vehicles\\Road\\Car - Volkswagen Golf (2007).v3d"
ModelFile2(0) = PTVVissimInstallationPath + "\\3DModels\\Vehicles\\Road\\HGV - EU 04 Tractor.v3d"
ModelFile2a = PTVVissimInstallationPath + "\\3DModels\\Vehicles\\Road\\HGV - EU 04a Trailer.v3d"
ModelFile2b = PTVVissimInstallationPath + "\\3DModels\\Vehicles\\Road\\HGV - EU 04b Trailer.v3d"
Model2D3DSegs(0) = ModelFile2(0)
Model2D3DSegs(1) = ModelFile2a
Vissim.Net.Models2D3D.AddModel2D3D 0, ModelFile1 ' unsigned int Key, SAFEARRAY(VARIANT) Model2D3DSegs (ModelFile1(String), ModelFile2, ...)
Set Model2D3D = Vissim.Net.Models2D3D.AddModel2D3D(0, Model2D3DSegs) ' unsigned int Key, SAFEARRAY(VARIANT) Model2D3DSegs (ModelFile1(String), ModelFile2, ...)
Model2D3D.Model2D3DSegs.AddModel2D3DSegment 0, ModelFile2b ' unsigned int Key, BSTR File3D
' Model2D3D.Model2D3DSegs.RemoveModel2D3DSegment Model2D3D.Model2D3DSegs.ItemByKey(1)

'--------------------------------------
' TimeIntervalSet
'--------------------------------------
' Possible values for the 'TimeIntervalSet' enumeration along with their integer representation:
'   AreaBehaviorType       8
'   ManagedLanesFacility   9
'   PedestrianInput        5
'   PedestrianRoutePartial 6
'   PedestrianRouteStatic 11
'   VehicleInput           1
'   VehicleRouteParking    4
'   VehicleRoutePartial    3
'   VehicleRoutePartialPT  7
'   VehicleRouteStatic     2
Dim TimeIntervalSet
TimeIntervalSet = 1 ' Enumeration 1 = VehicleInput
Vissim.Net.TimeIntervalSets.ItemByKey(TimeIntervalSet).TimeInts.AddTimeInterval(0) ' unsigned int Key

'--------------------------------------
' Links
'--------------------------------------
' Input parameters of AddLink:
' 1: unsigned int Key   = attribute Number (No)         | example: 0 or 123
' 2: BSTR WktLinestring = attribution Points3D          | example: "LINESTRING(PosX1 PosY1 PosZ1, PosX2 PosY2 PosZ2, ..., PosXn PosYn PosZn)" with Pos as double : PosZ is optional
' 3: LaneWidths         = number and widths of lanes    | example: (WidthLane1, WiidthLane2, ... WidthLaneN) as double
Dim LinksEnter(3)
Dim LaneWidth(0)
LaneWidth(0) = 3.5

Set LinksEnter(0) = Vissim.Net.Links.AddLink(0, "LINESTRING(-300 -2, -25 -2)", LaneWidth) ' unsigned int Key, BSTR WktLinestring "LINESTRING(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)", SAFEARRAY(VARIANT) LaneWidths (WidthLane1, WiidthLane2, ... WidthLaneN)
Set LinksEnter(1) = Vissim.Net.Links.AddLink(0, "LINESTRING(2 -300, 2 -25)", LaneWidth) ' unsigned int Key, BSTR WktLinestring "LINESTRING(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)", SAFEARRAY(VARIANT) LaneWidths (WidthLane1, WiidthLane2, ... WidthLaneN)
Set LinksEnter(2) = Vissim.Net.Links.AddLink(0, "LINESTRING(300 2, 25 2)", LaneWidth) ' unsigned int Key, BSTR WktLinestring "LINESTRING(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)", SAFEARRAY(VARIANT) LaneWidths (WidthLane1, WiidthLane2, ... WidthLaneN)
Set LinksEnter(3) = Vissim.Net.Links.AddLink(0, "LINESTRING(-2 300, -2 25)", LaneWidth) ' unsigned int Key, BSTR WktLinestring "LINESTRING(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)", SAFEARRAY(VARIANT) LaneWidths (WidthLane1, WiidthLane2, ... WidthLaneN)

' Opposite direction:
Dim LinksExit(3)
Dim cnt_Enter
for cnt_Enter = 0 to UBound(LinksEnter)
    Set LinksExit(cnt_Enter) = Vissim.Net.Links.GenerateOppositeDirection(LinksEnter(cnt_Enter), 1) 'ILink* Link, unsigned int NumberOfLanes
Next

'--------------------------------------
' Connectors
'--------------------------------------
' Input parameters of AddConnector
' 1: unsigned int Key   = attribute Number (No)         | example: 0 or 123
' 2: ILane* FromLane                                    | example: Vissim.Net.Links.ItemByKey(1).Lanes.ItemByKey(1)
' 3: double FromPos                                     | example: 200
' 4: ILane* ToLane                                      | example: Vissim.Net.Links.ItemByKey(2).Lanes.ItemByKey(1)
' 5: double ToPos                                       | example: 0
' 6: unsigned int NumberOfLanes                         | example: 1, 2, ...
' 7: BSTR WktLinestring = attribute Points3D            | example: "LINESTRING(PosX2 PosY2 PosZ2, ..., PosXn-1 PosYn-1 PosZn-1)" with Pos as double; PosZ is optional; Pos1 and Posn automatically created; No additional points: "LINESTRING EMPTY"
Dim cnt_Exit
for cnt_Enter = 0 to UBound(LinksEnter)
    for cnt_Exit = 0 to UBound(LinksExit)
        if cnt_Enter <> cnt_Exit then
            Vissim.Net.Links.AddConnector 0, LinksEnter(cnt_Enter).Lanes.ItemByKey(1), 275, LinksExit(cnt_Exit).Lanes.ItemByKey(1), 0, 1, "LINESTRING EMPTY" ' unsigned int Key, ILane* FromLane, double FromPos, ILane* ToLane, double ToPos, unsigned int NumberOfLanes, BSTR WktLinestring
		end if
	Next
Next

' alternative Route:
Dim LinksAlternative, ConnToAlternative, ConnFromAlternative
Set LinksAlternative = Vissim.Net.Links.AddLink(0, "LINESTRING(-150.000 -9.000 0.000,-113.009 -15.641 0.000,-75.582 -20.384 0.000,-37.864 -23.231 0.000,-0.000 -24.179 0.000,37.864 -23.231 0.000,75.582 -20.384 0.000,113.009 -15.641 0.000,150.000 -9.000 0.000)", LaneWidth) ' unsigned int Key, BSTR WktLinestring, SAFEARRAY(VARIANT) LaneWidths
Set ConnToAlternative = Vissim.Net.Links.AddConnector(0, LinksEnter(0).Lanes.ItemByKey(1), 130, LinksAlternative.Lanes.ItemByKey(1), 0, 1, "LINESTRING(-166.600 -4.181 0.000,-163.323 -5.284 0.000,-160.114 -6.775 0.000,-156.913 -8.368 0.000,-153.664 -9.779 0.000)") ' unsigned int Key, ILane* FromLane, double FromPos, ILane* ToLane, double ToPos, unsigned int NumberOfLanes, BSTR WktLinestring
Set ConnFromAlternative = Vissim.Net.Links.AddConnector(0, LinksAlternative.Lanes.ItemByKey(1), LinksAlternative.AttValue("Length2D"), LinksExit(2).Lanes.ItemByKey(1), 145, 1, "LINESTRING(153.741 -9.595 0.000,156.974 -7.762 0.000,160.112 -5.631 0.000,163.259 -3.606 0.000,166.520 -2.094 0.000)") ' unsigned int Key, ILane* FromLane, double FromPos, ILane* ToLane, double ToPos, unsigned int NumberOfLanes, BSTR WktLinestring
LinksAlternative.Lanes.AddLane 0, 3.5 ' unsigned int Key, double Width

'--------------------------------------
' Vehicle Inputs
'--------------------------------------
Dim VehInput
for cnt_Enter = 0 to UBound(LinksEnter)
    Set VehInput = Vissim.Net.VehicleInputs.AddVehicleInput(0, LinksEnter(cnt_Enter)) ' unsigned int Key, Link
    VehInput.AttValue("Volume(1)") = 500
Next

'--------------------------------------
' Parking Lot
'--------------------------------------
Dim ParkingLot1
Set ParkingLot1 = Vissim.Net.ParkingLots.AddParkingLot(1, LinksExit(2).Lanes.ItemByKey(1), 186) ' unsigned int Key, ILane* Lane, double Pos
Vissim.Net.Zones.AddZone(0) ' unsigned int Key
Vissim.Net.ParkingLots.AddAbstractParkingLot 0, LinksExit(3), 186, Vissim.Net.Zones.ItemByKey(1) ' unsigned int Key, ILink* Link, double Pos, IZone* Zone

'--------------------------------------
' Matrices
'--------------------------------------
Vissim.Net.Matrices.AddMatrix(0) ' unsigned int Key

'--------------------------------------
' Vehicle Routing Decisions
'--------------------------------------
Dim VehRoutDesSta, VehRoutDesPar, VehRoutDesPPT, VehRoutDesPark, VehRoutDesML, VehRoutDesDyn, VehRoutDesClose
Set VehRoutDesSta = Vissim.Net.VehicleRoutingDecisionsStatic.AddVehicleRoutingDecisionStatic(0, LinksEnter(0), 100) ' unsigned int Key, ILink* Link,, double Pos
Set VehRoutDesPar = Vissim.Net.VehicleRoutingDecisionsPartial.AddVehicleRoutingDecisionPartial(0, LinksEnter(0), 102, LinksExit(2), 202) ' unsigned int Key, ILink* Link, double Pos, ILink* DestLink, double DestPos
Set VehRoutDesPPT = Vissim.Net.VehicleRoutingDecisionsPartialPT.AddVehicleRoutingDecisionPartialPT(0, LinksEnter(0), 104, LinksExit(2), 204) ' unsigned int Key, ILink* Link, double Pos, ILink* DestLink, double DestPos
Set VehRoutDesPark = Vissim.Net.VehicleRoutingDecisionsParking.AddVehicleRoutingDecisionParking(0, LinksEnter(0), 106) ' unsigned int Key, ILink* Link, double Pos
Set VehRoutDesML = Vissim.Net.VehicleRoutingDecisionsManagedLanes.AddVehicleRoutingDecisionManagedLanes(0, LinksEnter(0), 108, LinksExit(2), 208) ' unsigned int Key, ILink* Link, double Pos, ILink* DestLink, double DestPos
Set VehRoutDesDyn = Vissim.Net.VehicleRoutingDecisionsDynamic.AddVehicleRoutingDecisionDynamic(0, LinksEnter(0), 110) ' unsigned int Key, ILink* Link, double Pos
Set VehRoutDesClose = Vissim.Net.VehicleRoutingDecisionsClosure.AddVehicleRoutingDecisionClosure(0, LinksEnter(0), 112) ' unsigned int Key, ILink* Link, double Pos

'--------------------------------------
' Vehicle Routes
'--------------------------------------
Dim VehRoutSta1, VehRoutSta2, RouteSegment1(1), RouteSegment2(5)
Set RouteSegment1(0) = LinksAlternative
RouteSegment1(1) = 20
Set RouteSegment2(0) = ConnToAlternative
RouteSegment2(1) = 1
Set RouteSegment2(2) = LinksAlternative
RouteSegment2(3) = 20
Set RouteSegment2(4) = ConnFromAlternative
RouteSegment2(5) = 1
Set VehRoutSta1 = VehRoutDesSta.VehRoutSta.AddVehicleRouteStatic(0, LinksExit(2), 200) ' unsigned int Key, ILink* DestLink, double DestPos
Set VehRoutSta2 = VehRoutDesSta.VehRoutSta.AddVehicleRouteStatic(0, LinksExit(2), 200) ' unsigned int Key, ILink* DestLink, double DestPos
VehRoutSta2.UpdateLinkSequence(RouteSegment1) ' SAFEARRAY(VARIANT) RouteSegments Recomputes the link sequence. RouteSegments must be a list of each link and its offset that should be part of the route.
VehRoutSta2.UpdateLinkSequence(RouteSegment2) ' SAFEARRAY(VARIANT) RouteSegments

VehRoutDesPar.VehRoutPart.AddVehicleRoutePartial(0) ' unsigned int Key
VehRoutDesPPT.VehRoutPartPT.AddVehicleRoutePartialPT(0) ' unsigned int Key
VehRoutDesPark.VehRoutPark.AddVehicleRouteParking 0, ParkingLot1 ' unsigned int Key, IParkingLot* ParkLot
VehRoutDesML.VehRoutMngLns.AddVehicleRouteManagedLanes 0, 1 ' unsigned int Key, enum ManagedLaneType Type (1: Managed; 2: General purpose)
VehRoutDesClose.VehRoutClos.AddVehicleRouteClosure 0, LinksExit(2), 202 ' unsigned int Key, ILink* DestLink, double DestPos

Dim Lane1, Lane2, Lane3, Lane4
Set Lane1 = LinksEnter(0).Lanes.ItemByKey(1)
Set Lane2 = LinksEnter(1).Lanes.ItemByKey(1)
Set Lane3 = LinksEnter(2).Lanes.ItemByKey(1)
Set Lane4 = LinksEnter(3).Lanes.ItemByKey(1)

'--------------------------------------
' Stop Sign
'--------------------------------------
Vissim.Net.StopSigns.AddStopSign 0, Lane1, 274 ' unsigned int Key, ILane* Lane, double Pos

'--------------------------------------
' Priority Rule
'--------------------------------------
Dim PriRule, ConflictMarkers(3)
Set ConflictMarkers(0) = Lane2
ConflictMarkers(1) = 270
Set ConflictMarkers(2) = Lane3
ConflictMarkers(3) = 270
Set PriRule = Vissim.Net.PriorityRules.AddPriorityRule(0, Lane1, 270, ConflictMarkers) ' unsigned int Key, ILane* Lane, double Pos, SAFEARRAY(VARIANT) ConflictMarkers (ConflictMarkers must be a list that specifies for each ConflictMarker alternating its lane and position)
PriRule.ConflictMarkers.AddConflictMarker 0, Lane4, 270 'unsigned int Key, ILane* Lane, double Pos

'--------------------------------------
' Desired Speed Decision
'--------------------------------------
Vissim.Net.DesSpeedDecisions.AddDesSpeedDecision 0, Lane1, 10 ' int Key, ILane* Lane, double Pos

'--------------------------------------
' Reduced Speed Area
'--------------------------------------
Vissim.Net.ReducedSpeedAreas.AddReducedSpeedArea 0, Lane1, 10 ' unsigned int Key, ILane* Lane, double Pos

'--------------------------------------
' Node
'--------------------------------------
Vissim.Net.Nodes.AddNode 0, "POLYGON((-30 -30, 30 -30, 30 30, -30 30, -30 -30))" ' unsigned int Key, BSTR WktPolygon

'--------------------------------------
' Public Transport
'--------------------------------------
Dim PTLine, RouteSegment3(1)
Set RouteSegment3(0) = LinksAlternative
RouteSegment3(1) = 20
Vissim.Net.PTStops.AddPTStop 0, Lane1, 150 ' unsigned int Key, ILane* Lane, double Pos
Set PTLine = Vissim.Net.PTLines.AddPTLine(0, LinksEnter(0), LinksExit(2), 200) ' unsigned int Key, ILink* EntryLink, ILink* DestLink, double DestPos
PTLine.DepTimes.AddDepartureTime 0, 123 ' unsigned int Key, double Dep
PTLine.DepTimes.AddDepartureTime 0, 246 ' unsigned int Key, double Dep
PTLine.UpdateLinkSequence(RouteSegment3) ' SAFEARRAY(VARIANT) RouteSegments Recomputes the link sequence. RouteSegments must be a list of each link and its offset that should be part of the route.

'--------------------------------------
' Static 3D Models
'--------------------------------------
Dim Static3DModelFile
Static3DModelFile = PTVVissimInstallationPath + "\\3DModels\\Vegetation\\Tree 04.v3d"
Vissim.Net.Static3DModels.AddStatic3DModel 0, Static3DModelFile, "Point(-100, -100, 0)" ' unsigned int Key, BSTR ModelFilename, BSTR PosWktPoint3D

'--------------------------------------
' Background Image
'--------------------------------------
Dim BackgroundImageFile
BackgroundImageFile = PTVVissimInstallationPath + "\\3DModels\\Textures\\Marsh.png"
Vissim.Net.BackgroundImages.AddBackgroundImage 0, BackgroundImageFile, "Point(100, 100)", "Point(-100, -100)" ' unsigned int Key, BSTR PathFilename, BSTR PosTRWktPoint, BSTR PosBLWktPoint

'--------------------------------------
' Presentation
'--------------------------------------
Dim Storyboard, camPos
Set camPos = Vissim.Net.CameraPositions.AddCameraPosition (0, "Point(-100, -100, 100)")  ' unsigned int Key, BSTR WktPoint3D
camPos.AttValue("Name") = "Position 1 (added by COM)"
Set Storyboard = Vissim.Net.Storyboards.AddStoryboard(0) ' unsigned int Key
Storyboard.Keyframes.AddKeyframe(0) ' unsigned int Key

'--------------------------------------
' Section
'--------------------------------------
Vissim.Net.Sections.AddSection 0, "POLYGON((-200 -150,  200 -150,  200 150,  -200 150))" ' unsigned int Key, BSTR WktPolygon

'--------------------------------------
' Signal Controller
'--------------------------------------
Dim SignalController
Set SignalController = Vissim.Net.SignalControllers.AddSignalController(0) ' unsigned int Key
SignalController.SGs.AddSignalGroup(0) ' unsigned int Key

'--------------------------------------
' Detector
'--------------------------------------
Vissim.Net.Detectors.AddDetector 0, Lane1, 250 ' unsigned int Key, Lane, double Pos

'--------------------------------------
' Signal Head
'--------------------------------------
Vissim.Net.SignalHeads.AddSignalHead 0, Lane1, 274 ' unsigned int Key, ILane* Lane, double Pos

'--------------------------------------
' 3D Traffic Signal
'--------------------------------------
Dim TrafficSignal3D
Set TrafficSignal3D = Vissim.Net.Signals3D.AddTrafficSignal3D(0, "POINT(0 0)") ' unsigned int Key, WktPos3D As String
TrafficSignal3D.SigHeads.AddSignalHead3D(0) ' unsigned int Key
TrafficSignal3D.SigArms.AddSignalArm3D(0) ' unsigned int Key
TrafficSignal3D.TrafficSigns.AddTrafficSign3D(0) ' unsigned int Key
TrafficSignal3D.Streetlights.AddStreetlight3D(0) ' unsigned int Key

'--------------------------------------
' Pavement Marking
'--------------------------------------
Vissim.Net.PavementMarkings.AddPavementMarking 0, Lane1, 215 ' unsigned int Key, ILane* Lane, double Pos

'--------------------------------------
' Vehicle Evaluations
'--------------------------------------
Vissim.Net.DataCollectionPoints.AddDataCollectionPoint 0, Lane1, 230 ' unsigned int Key, ILane* Lane, double Pos
Vissim.Net.VehicleTravelTimeMeasurements.AddVehicleTravelTimeMeasurement 0, LinksEnter(1), 50, LinksExit(0), 200 ' unsigned int Key, ILink* StartLink, double StartPos, ILink* EndLink, double EndPos
Vissim.Net.QueueCounters.AddQueueCounter 0, LinksEnter(0), 123 ' unsigned int Key, ILink* Link, double Pos
Vissim.Net.DelayMeasurements.AddDelayMeasurement 0 ' unsigned int Key




'=============================================
'=============================================
' ADDING OBJECTS SPECIFICALLY FOR PEDESTRIANS
'=============================================
'=============================================

'--------------------------------------
' Pedestrian Base Data
'--------------------------------------
Dim AreaBehaviorType
Vissim.Net.PedestrianTypes.AddPedestrianType(0) ' unsigned int Key
Vissim.Net.PedestrianClasses.AddPedestrianClass(0) ' unsigned int Key
Vissim.Net.WalkingBehaviors.AddWalkingBehavior(0) ' unsigned int Key
Set AreaBehaviorType = Vissim.Net.AreaBehaviorTypes.AddAreaBehaviorType(0) ' unsigned int Key
AreaBehaviorType.AreaBehavTypeElements.AddAreaBehaviorTypeElement(Vissim.Net.PedestrianClasses.ItemByKey(30)) ' IPedestrianClass* PedClass



'--------------------------------------
' Pedestrian Compositions
'--------------------------------------
Dim PedCompRelFlow1(3), PedComposition
Set PedCompRelFlow1(0) = Vissim.Net.PedestrianTypes.ItemByKey(100)
Set PedCompRelFlow1(1) = Vissim.Net.DesSpeedDistributions.ItemByKey(5)
Set PedCompRelFlow1(2) = Vissim.Net.PedestrianTypes.ItemByKey(200)
Set PedCompRelFlow1(3) = Vissim.Net.DesSpeedDistributions.ItemByKey(5)
Set PedComposition = Vissim.Net.PedestrianCompositions.AddPedestrianComposition(0, PedCompRelFlow1) ' unsigned int Key, SAFEARRAY(VARIANT) PedCompRelFlows
PedComposition.PedCompRelFlows.AddPedestrianCompositionRelativeFlow Vissim.Net.PedestrianTypes.ItemByKey(300), Vissim.Net.DesSpeedDistributions.ItemByKey(5) ' IPedestrianType* PedType, IDesSpeedDistribution* DesSpeedDistr

'-----------------------
' Areas
'-----------------------
' Input parameters of AddArea:
' 1: unsigned int Key   = attribute Number (No)          | example: 0 or 123
' 2: BSTR WktPolygon    = attribute Points               | example: "POLYGON(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)" with Pos as double

' Example: Area1 = Vissim.Net.Areas.AddArea(0, "POLYGON((0 50, 50 50, 50 100, 0 100))") ' unsigned int Key, BSTR WktPolygon

' Create some areas:
Dim xydim(1), StartPosAll(9,1), Area(11), cnt, x1, x2, y1, y2, Polyg
xydim(0) = 10
xydim(1) = 10
'StartPosAll(0, 0) = ((50, 50), (50, 60), (50, 70), (50, 80), (50, 90), (60, 60), (70, 60), (70, 70), (70, 80), (60, 80))
StartPosAll(0, 0) = 50
StartPosAll(0, 1) = 50
StartPosAll(1, 0) = 50
StartPosAll(1, 1) = 60
StartPosAll(2, 0) = 50
StartPosAll(2, 1) = 70
StartPosAll(3, 0) = 50
StartPosAll(3, 1) = 80
StartPosAll(4, 0) = 50
StartPosAll(4, 1) = 90
StartPosAll(5, 0) = 60
StartPosAll(5, 1) = 60
StartPosAll(6, 0) = 70
StartPosAll(6, 1) = 60
StartPosAll(7, 0) = 70
StartPosAll(7, 1) = 70
StartPosAll(8, 0) = 70
StartPosAll(8, 1) = 80
StartPosAll(9, 0) = 60
StartPosAll(9, 1) = 80
for cnt = 0 to UBound(StartPosAll)
    x1 = Cstr(StartPosAll(cnt, 0))
    x2 = Cstr(StartPosAll(cnt, 0) + xydim(0))
    y1 = Cstr(StartPosAll(cnt, 1))
    y2 = Cstr(StartPosAll(cnt, 1) + xydim(1))
    Polyg = "POLYGON((" + x1 + " " + y1 +"," + x1 + " " + y2 + "," + x2 + " " + y2 + "," + x2 + " " + y1 + "))"
    Set Area(cnt) = Vissim.Net.Areas.AddArea(0, Polyg) ' unsigned int Key, BSTR WktPolygon
Next

'--------------------------------------
' Obstacle & Ramp
'--------------------------------------
Dim Ramp
Vissim.Net.Obstacles.AddObstacle 0, "POLYGON((55 72.5, 65 72.5, 65 77.5, 55 77.5))" ' unsigned int Key, BSTR WktPolygon
Set Ramp = Vissim.Net.Ramps.AddRamp(0, "POLYGON((50 100,  50 110,  60 110,  60 100))") ' unsigned int Key, BSTR WktPolygon
Set Area(cnt + 1) = Vissim.Net.Areas.AddArea(0, "POLYGON((50 110,  60 110,  60 120,  50 120))") ' unsigned int Key, BSTR WktPolygon

'--------------------------------------
' Pedestrian Input
'--------------------------------------
Vissim.Net.PedestrianInputs.AddPedestrianInput 0, Area(0) ' unsigned int Key, IArea* Area

'--------------------------------------
' Pedestrian Routes
'--------------------------------------
Dim PedRouteDesSta1, PedRouteSta1, PedRouteSta2, PedRouteSta3, PedRouteDesPart1, PedRoutePart1, PedRoutePart2, PedRouteDesPart2
Set PedRouteDesSta1 = Vissim.Net.PedestrianRoutingDecisionsStatic.AddPedestrianRoutingDecisionStatic(0, Area(0)) ' unsigned int Key, IArea* Area
Set PedRouteSta1 = PedRouteDesSta1.PedRoutSta.AddPedestrianRouteStaticOnArea(0, Area(1)) ' unsigned int Key, IArea* Area
PedRouteSta1.PedRoutLoc.AddPedestrianRouteLocationOnArea 0, Area(2) ' unsigned int Key, IArea* Area | => will become next route end
PedRouteSta1.PedRoutLoc.AddPedestrianRouteLocationOnArea 0, Area(5) ' unsigned int Key, IArea* Area | => will become next route end
PedRouteSta1.PedRoutLoc.AddPedestrianRouteLocationOnArea 0, Area(3) ' unsigned int Key, IArea* Area | => will become next route end
PedRouteSta1.PedRoutLoc.AddPedestrianRouteLocationOnArea 0, Area(4) ' unsigned int Key, IArea* Area | => will become next route end

Set PedRouteSta2 = PedRouteDesSta1.PedRoutSta.AddPedestrianRouteStaticOnRamp(0, Ramp)  ' unsigned int Key, IRamp* PedRoutLoc
Set PedRouteSta3 = PedRouteDesSta1.PedRoutSta.AddPedestrianRouteStaticOnArea(0, Area(9)) ' unsigned int Key, IArea* Area
PedRouteSta3.PedRoutLoc.AddPedestrianRouteLocationOnRamp 0, Ramp ' unsigned int Key, IRamp* PedRoutLoc | => will become next route end

Set PedRouteDesPart1 = Vissim.Net.PedestrianRoutingDecisionsPartial.AddPedestrianRoutingDecisionPartial(0, Area(1)) ' unsigned int Key, IArea* Area
Set PedRoutePart1 = PedRouteDesPart1.PedRoutPart.AddPedestrianRoutePartialOnArea(0, Area(7)) ' unsigned int Key, IArea* Area
PedRoutePart1.PedRoutLoc.AddPedestrianRouteLocationOnArea 0, Area(3) ' unsigned int Key, IArea* Area | => will become next route end
Set PedRoutePart2 = PedRouteDesPart1.PedRoutPart.AddPedestrianRoutePartialOnArea(0, Area(3)) ' unsigned int Key, IArea* Area

Set PedRouteDesPart2 = Vissim.Net.PedestrianRoutingDecisionsPartial.AddPedestrianRoutingDecisionPartial(0, Area(1)) ' unsigned int Key, IArea* Area
PedRouteDesPart2.PedRoutPart.AddPedestrianRoutePartialOnRamp 0, Ramp  ' unsigned int Key, IRamp* PedRoutLoc

'--------------------------------------
' Pedestrian Evaluation
'--------------------------------------
Vissim.Net.PedestrianTravelTimeMeasurements.AddPedestrianTravelTimeMeasurement 0, Area(0), Area(4) ' unsigned int Key, IArea* StartArea, IArea* EndArea

'======================================
' END OF SCRIPT
'======================================















