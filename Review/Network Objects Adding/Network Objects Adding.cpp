// Network Objects Adding.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "Network Objects Adding.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#include <iostream>
#include<conio.h>
#include <objbase.h>
#include <vector>
#include <atlstr.h>             // CString
#include <atlcomcli.h>          // CComVariant
#include <atlsafe.h>            // CComSafeArray
#include <iostream>     // cout, ios
#include <sstream>      // ostringstream

// import the *.exe file providing the COM interface
// use your Vissim installation path here as the example shows
#import "c:\\Program Files\\PTV Vision\\PTV Vissim 22\\Exe\\VISSIM\\VISSIM220.exe"


using namespace VISSIMLIB;
using namespace std;

wstring variant2String(_variant_t v)
{
  try
  {
    v.ChangeType(VT_BSTR);
    wstring s = v.bstrVal;
    return s;
  }
  catch (_com_error e)
  {
    return L"";
  }
}

void makeSafeArray(CComSafeArray<VARIANT> &v, int idx) // root of the recursive variadic function to add the next argument to the safearray
{
}

template<typename T, typename... Targs>
void makeSafeArray(CComSafeArray<VARIANT> &v, int idx, T value, Targs... Fargs) // recursive variadic function to add the next argument to the safearray
{
  v[idx++] = value;
  makeSafeArray(v, idx, Fargs...);
}

template<typename... Targs>
_variant_t makeVariantSafeArray(Targs... values) // variadic function to make a safearray from the given arguments
{
  auto count = sizeof...(values);
  if (count > 0) {
    CComSafeArray<VARIANT> v(count);
    makeSafeArray(v, 0, values...);
    variant_t retVar;
    retVar.vt = VT_VARIANT | VT_ARRAY;
    retVar.parray = *v.GetSafeArrayPtr();
    return retVar;
  }
  else {
    CComSafeArray<VARIANT> v((ULONG)0);
    variant_t retVar;
    retVar.vt = VT_VARIANT | VT_ARRAY;
    retVar.parray = *v.GetSafeArrayPtr();
    return retVar;
  }
}



// The one and only application object

CWinApp theApp;

using namespace std;

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int nRetCode = 0;

	HMODULE hModule = ::GetModuleHandle(NULL);

	if (hModule != NULL) {
		// initialize MFC and print and error on failure
		if (!AfxWinInit(hModule, NULL, ::GetCommandLine(), 0)) {
			// TODO: change error code to suit your needs
			_tprintf(_T("Fatal Error: MFC initialization failed\n"));
			nRetCode = 1;
		}
		else {
      //==========================================================================
      // C++ - Script for Vissim 9 +
      // Copyright(C) PTV AG.All rights reserved.
      // Stefan Hengst 2017
      // -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
      // Example for building a network through Vissim COM
      //==========================================================================

      // This script demonstrates how to add and remove network objects in C++ (with early binding).
      // All "Add' methods are descried in the PTV Vissim COM help, which can be accessed in the PTV Vissim menu "Help - COM Help'.
      // If you search for "add*' in the COM Help, your find all "Add' methods.
      // For all objects where the key is the "Number' attribute, the value of 0 means that Vissim will automatically choose a number that is not in use.

      // There are also methods for removing objects.You can find the remove methods if you search for "remove*' in the COM help.
      // Example to remove a link : pVissim->GetNet()->GetLinks()->RemoveLink(pVissim->GetNet()->GetLinks()->GetItemByKey(1))

      // initialize COM
      CoInitializeEx(NULL, COINIT_MULTITHREADED);

      // COM - Server
      // Connecting the COM Server
      IVissimPtr pVissim;
      HRESULT hr = pVissim.CreateInstance("Vissim.Vissim");
	  // Specific version: HRESULT hr = pVissim.CreateInstance("Vissim.Vissim.100");
      if (hr != S_OK) {
        cout << "COM connection to Vissim cannot be established." << endl;
      }
      else {
        try {

          // fetch some general information about Vissim
          wstring PTVVissimInstallationPath = variant2String(pVissim->GetAttValue((bstr_t)"ExeFolder"));    // directory of your PTV Vissim installation(where vissim.exe is located)
          wstring version = variant2String(pVissim->GetAttValue((bstr_t)"VERSION"));
          wstring versionNumber = variant2String(pVissim->GetAttValue((bstr_t)"VERSIONNUMBER"));
          wstring aName = variant2String(pVissim->GetAttValue((bstr_t)"APPLICATIONNAME"));
          wstring wTitle = variant2String(pVissim->GetAttValue((bstr_t)"WINDOWTITLE"));
          wstring rev = variant2String(pVissim->GetAttValue((bstr_t)"REVISION"));
          wprintf(L"%s\n%s\n%s\n%s\n%s\n\n%s\n", version.c_str(), versionNumber.c_str(), aName.c_str(), wTitle.c_str(), rev.c_str(), PTVVissimInstallationPath.c_str());

          // start new network
          pVissim->New();

          // For this example, units are set to Metric.
          // Note: PTV Vissim coordinates are always in meters[m]
          _variant_t UnitCurrent = pVissim->GetNet()->GetNetPara()->GetAttValue((bstr_t)L"UnitLenShort");
          vector<wstring> UnitAttributes{ L"UnitAccel", L"UnitLenLong", L"UnitLenShort", L"UnitLenVeryShort", L"UnitSpeed", L"UnitSpeedSmall" };
          for (wstring UnitAttrCurr : UnitAttributes) {
            pVissim->GetNet()->GetNetPara()->PutAttValue((bstr_t)UnitAttrCurr.c_str(), (int)0);         // 0: Metric[m], 1 : Imperial[ft]
          }

          // Zoom Network Editor
          pVissim->GetGraphics()->GetCurrentNetworkWindow()->ZoomTo(-300, -300, 300, 300);

          //======================================
          //======================================
          // ADDING OBJECTS FOR VEHICULAR TRAFFIC
          //======================================
          //======================================

          //--------------------------------------
          // Base Data
          //--------------------------------------
          pVissim->GetNet()->GetVehicleClasses()->AddVehicleClass(0); // unsigned int Key
          pVissim->GetNet()->GetVehicleTypes()->AddVehicleType(0); // unsigned int Key
          pVissim->GetNet()->GetDrivingBehaviors()->AddDrivingBehavior(0); // unsigned int Key
          pVissim->GetNet()->GetLinkBehaviorTypes()->AddLinkBehaviorType(0); // unsigned int Key
          pVissim->GetNet()->GetLinkBehaviorTypes()->GetItemByKey(1)->GetVehClassDrivBehav()->AddVehClassDrivingBehavior(pVissim->GetNet()->GetVehicleClasses()->GetItemByKey(30)); // IVehicleClass* VehClass

          pVissim->GetNet()->GetDisplayTypes()->AddDisplayType(0); // unsigned int Key
          pVissim->GetNet()->GetLevels()->AddLevel(0); // unsigned int Key

          //--------------------------------------
          // Vehicle Compositions
          //--------------------------------------
          pVissim->GetNet()->GetVehicleCompositions()->AddVehicleComposition(0
            , makeVariantSafeArray()
            ); // unsigned int Key, SAFEARRAY(VARIANT) VehCompRelFlows
          pVissim->GetNet()->GetVehicleCompositions()->AddVehicleComposition(0
            , makeVariantSafeArray(pVissim->GetNet()->GetVehicleTypes()->GetItemByKey(100).GetInterfacePtr(), pVissim->GetNet()->GetDesSpeedDistributions()->GetItemByKey(40).GetInterfacePtr())
            ); // unsigned int Key, SAFEARRAY(VARIANT) VehCompRelFlows
          pVissim->GetNet()->GetVehicleCompositions()->AddVehicleComposition(9
            , makeVariantSafeArray(pVissim->GetNet()->GetVehicleTypes()->GetItemByKey(100).GetInterfacePtr(), pVissim->GetNet()->GetDesSpeedDistributions()->GetItemByKey(40).GetInterfacePtr(), pVissim->GetNet()->GetVehicleTypes()->GetItemByKey(200).GetInterfacePtr(), pVissim->GetNet()->GetDesSpeedDistributions()->GetItemByKey(30).GetInterfacePtr())
            ); // unsigned int Key, SAFEARRAY(VARIANT) VehCompRelFlows
          pVissim->GetNet()->GetVehicleCompositions()->GetItemByKey(9)->GetVehCompRelFlows()->AddVehicleCompositionRelativeFlow(pVissim->GetNet()->GetVehicleTypes()->GetItemByKey(300)
            , pVissim->GetNet()->GetDesSpeedDistributions()->GetItemByKey(25)); // IVehicleType* VehType, IDesSpeedDistribution* DesSpeedDistr

          //--------------------------------------
          // Functions
          //--------------------------------------
          pVissim->GetNet()->GetDesAccelerationFunctions()->AddDesAccelerationFunction(0
            , makeVariantSafeArray(0, 3, 2, 4, 120, 6, 5, 7, 160, 0, 0, 0)); // unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts[X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...]
          pVissim->GetNet()->GetDesDecelerationFunctions()->AddDesDecelerationFunction(0
            , makeVariantSafeArray(0, -3, -4, -2, 120, -6, -7, -5, 160, 0, 0, 0)); // unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts[X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...]
          pVissim->GetNet()->GetMaxAccelerationFunctions()->AddMaxAccelerationFunction(0
            , makeVariantSafeArray(0, 3, 2, 4, 120, 6, 5, 7, 160, 0, 0, 0)); // unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts[X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...]
          pVissim->GetNet()->GetMaxDecelerationFunctions()->AddMaxDecelerationFunction(0
            , makeVariantSafeArray(0, -3, -4, -2, 120, -6, -7, -5, 160, 0, 0, 0)); // unsigned int Key, SAFEARRAY(VARIANT) AccelFuncDataPts[X1(Speed), Y1(Acceleration), YMin1(MinAcceleration), YMax1(MaxAcceleration), X2, Y2, YMin2, YMax2, ...]

          //--------------------------------------
          // Distributions
          //--------------------------------------
          pVissim->GetNet()->GetDesSpeedDistributions()->AddDesSpeedDistribution(0
            , makeVariantSafeArray(50, 0, 60, 0.1, 70, 0.9, 80, 1)); // unsigned int Key, SAFEARRAY(VARIANT) SpeedDistrDatPts[X1(speed), Y1(0...1), X2, Y2, ...]
          pVissim->GetNet()->GetPowerDistributions()->AddPowerDistribution(0
            , makeVariantSafeArray(60, 0, 75, 0.5, 100, 1)); // unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts[X1(power), Y1(0...1), X2, Y2, ...]
          pVissim->GetNet()->GetWeightDistributions()->AddWeightDistribution(0
            , makeVariantSafeArray(500, 0, 1000, 0.5, 2000, 1)); // unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts[X1(weight), Y1(0...1), X2, Y2, ...]
          pVissim->GetNet()->GetTimeDistributions()->AddTimeDistributionNormal(0); // unsigned int Key
          pVissim->GetNet()->GetTimeDistributions()->AddTimeDistributionEmpirical(0
            , makeVariantSafeArray(10, 0, 12, 0.75, 15, 1)); // unsigned int Key, SAFEARRAY(VARIANT) DurDistrDataPts[X1(duration), Y1(0...1), X2, Y2, ...]
          pVissim->GetNet()->GetLocationDistributions()->AddLocationDistribution(0
            , makeVariantSafeArray(0, 0, 0.5, 0.3, 1, 1)); // unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts[X1(location), Y1(0...1), X2, Y2, ...]
          pVissim->GetNet()->GetDistanceDistributions()->AddDistanceDistribution(0
            , makeVariantSafeArray(50, 0, 90, 0.2, 110, 0.8, 150, 1)); // unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts[X1(Distance), Y1(0...1), X2, Y2, ...]
          pVissim->GetNet()->GetOccupancyDistributions()->AddOccupancyDistributionNormal(0); // unsigned int Key
          pVissim->GetNet()->GetOccupancyDistributions()->AddOccupancyDistributionEmpirical(0
            , makeVariantSafeArray(1, 0, 2, 0.66, 3, 0.8, 5, 0.9, 9, 1)); // unsigned int Key, SAFEARRAY(VARIANT) DistrDataPts[X1(Distance), Y1(0...1), X2, Y2, ...]
          // Replace of all Distribution points(does only work with empirical distributions) :
          pVissim->GetNet()->GetDesSpeedDistributions()->GetItemByKey(5)->GetSpeedDistrDatPts()->ReplaceAll(makeVariantSafeArray(3.9, 0, 4.1, 1)); // SAFEARRAY(VARIANT) Elements

          IColorDistributionPtr ColorDistribution = pVissim->GetNet()->GetColorDistributions()->AddColorDistribution(0
            , makeVariantSafeArray(L"44556677", L"ff000000")); // unsigned int Key, SAFEARRAY(VARIANT) ColorDistrEl[Color1(StringHEX), Color2, ...]
          ColorDistribution->GetColorDistrEl()->AddColorDistributionElement(0, L"ff994400"); // unsigned int Key, BSTR Color(StringHEX)
          ColorDistribution->GetColorDistrEl()->ReplaceAll(makeVariantSafeArray(L"11223344", L"ffffffff")); // SAFEARRAY(VARIANT) ColorDistrEl[Color1(StringHEX), Color2, ...]

          //--------------------------------------
          // 2D / 3D Models
          //--------------------------------------
          // Model2D3D.Model2D3DSegs.AddModel2D3DSegment(0, "C:\\1.v3d') // unsigned int Key, BSTR File3D
          wstring ModelFile1 = PTVVissimInstallationPath + L"\\3DModels\\Vehicles\\Road\\Car - Volkswagen Golf (2007).v3d";
          wstring ModelFile2 = PTVVissimInstallationPath + L"\\3DModels\\Vehicles\\Road\\HGV - EU 04 Tractor.v3d";
          wstring ModelFile2a = PTVVissimInstallationPath + L"\\3DModels\\Vehicles\\Road\\HGV - EU 04a Trailer.v3d";
          wstring ModelFile2b = PTVVissimInstallationPath + L"\\3DModels\\Vehicles\\Road\\HGV - EU 04b Trailer.v3d";
          pVissim->GetNet()->GetModels2D3D()->AddModel2D3D(0, makeVariantSafeArray(ModelFile1.c_str())); // unsigned int Key, SAFEARRAY(VARIANT) Model2D3DSegs[ModelFile1(String), ModelFile2, ...]
          IModel2D3DPtr Model2D3D = pVissim->GetNet()->GetModels2D3D()->AddModel2D3D(0
            , makeVariantSafeArray(ModelFile2.c_str(), ModelFile2a.c_str())); // unsigned int Key, SAFEARRAY(VARIANT) Model2D3DSegs[ModelFile1(String), ModelFile2, ...]
          Model2D3D->GetModel2D3DSegs()->AddModel2D3DSegment(0, ModelFile2b.c_str()); // unsigned int Key, BSTR File3D
          // Model2D3D.Model2D3DSegs.RemoveModel2D3DSegment(Model2D3D.Model2D3DSegs.GetItemByKey(1))

          IModel2D3DDistributionPtr Model2D3DDistribution = pVissim->GetNet()->GetModel2D3DDistributions()->AddModel2D3DDistribution(0
            , makeVariantSafeArray(pVissim->GetNet()->GetModels2D3D()->GetItemByKey(1).GetInterfacePtr(), pVissim->GetNet()->GetModels2D3D()->GetItemByKey(2).GetInterfacePtr())); // unsigned int Key, SAFEARRAY(VARIANT) Model2D3DDistrEl[Model1(IModel2D3D), Model2, ...]
          Model2D3DDistribution->GetModel2D3DDistrEl()->AddModel2D3DDistributionElement(0
            , pVissim->GetNet()->GetModels2D3D()->GetItemByKey(3)); // unsigned int Key, IModel2D3D* Model2D3D
          Model2D3DDistribution->GetModel2D3DDistrEl()->ReplaceAll(
            makeVariantSafeArray(pVissim->GetNet()->GetModels2D3D()->GetItemByKey(5).GetInterfacePtr(), pVissim->GetNet()->GetModels2D3D()->GetItemByKey(6).GetInterfacePtr())); // SAFEARRAY(VARIANT) Model2D3DDistrEl[Model1(IModel2D3D), Model2, ...]

          //--------------------------------------
          // TimeIntervalSet
          //--------------------------------------
          // Possible values for the "TimeIntervalSet' enumeration along with their integer representation :
          //   AreaBehaviorType       8
          //   ManagedLanesFacility   9
          //   PedestrianInput        5
          //   PedestrianRoutePartial 6
          //   PedestrianRouteStatic 11
          //   VehicleInput           1
          //   VehicleRouteParking    4
          //   VehicleRoutePartial    3
          //   VehicleRoutePartialPT  7
          //   VehicleRouteStatic     2
          int TimeIntervalSet = 1;   // 1 = VehicleInput
          pVissim->GetNet()->GetTimeIntervalSets()->GetItemByKey(TimeIntervalSet)->GetTimeInts()->AddTimeInterval(0); // unsigned int Key

          //--------------------------------------
          // Links
          //--------------------------------------
          // Input parameters of AddLink :
          // 1: unsigned int Key = attribute Number(No) | example : 0 or 123
          // 2: BSTR WktLinestring = attribution Points3D | example : "LINESTRING(PosX1 PosY1 PosZ1, PosX2 PosY2 PosZ2, ..., PosXn PosYn PosZn)' with Pos as double : PosZ is optional
          // 3: LaneWidths = number and widths of lanes | example : [WidthLane1, WiidthLane2, ... WidthLaneN] as double
          vector<ILinkPtr> LinksEnter(4);
          LinksEnter[0] = pVissim->GetNet()->GetLinks()->AddLink(0, L"LINESTRING(-300 -2, -25 -2)", makeVariantSafeArray(3.5)); // unsigned int Key, BSTR WktLinestring "LINESTRING(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)', SAFEARRAY(double) LaneWidths[WidthLane1, WiidthLane2, ... WidthLaneN]
          LinksEnter[1] = pVissim->GetNet()->GetLinks()->AddLink(0, L"LINESTRING(2 -300, 2 -25)", makeVariantSafeArray(3.5)); // unsigned int Key, BSTR WktLinestring "LINESTRING(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)', SAFEARRAY(double) LaneWidths[WidthLane1, WiidthLane2, ... WidthLaneN]
          LinksEnter[2] = pVissim->GetNet()->GetLinks()->AddLink(0, L"LINESTRING(300 2, 25 2)", makeVariantSafeArray(3.5)); // unsigned int Key, BSTR WktLinestring "LINESTRING(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)', SAFEARRAY(double) LaneWidths[WidthLane1, WiidthLane2, ... WidthLaneN]
          LinksEnter[3] = pVissim->GetNet()->GetLinks()->AddLink(0, L"LINESTRING(-2 300, -2 25)", makeVariantSafeArray(3.5)); // unsigned int Key, BSTR WktLinestring "LINESTRING(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)', SAFEARRAY(double) LaneWidths[WidthLane1, WiidthLane2, ... WidthLaneN]

          // Opposite direction :
          vector<ILinkPtr> LinksExit(4);
          for (vector<ILinkPtr>::size_type cnt_Enter = 0; cnt_Enter < LinksEnter.size(); ++cnt_Enter) {
            LinksExit[cnt_Enter] = pVissim->GetNet()->GetLinks()->GenerateOppositeDirection(LinksEnter[cnt_Enter], 1); //ILink* Link, unsigned int NumberOfLanes
          }

          //--------------------------------------
          // Connectors
          //--------------------------------------
          // Input parameters of AddConnector :
          // 1: unsigned int Key = attribute Number(No) | example : 0 or 123
          // 2: ILane* FromLane | example : pVissim->GetNet()->GetLinks()GetItemByKey(1)->GetLanes()->GetItemByKey(1)
          // 3: double FromPos | example : 200
          // 4: ILane* ToLane | example : pVissim->GetNet()->GetLinks()GetItemByKey(2)->GetLanes()->GetItemByKey(1)
          // 5: double ToPos | example : 0
          // 6: unsigned int NumberOfLanes | example : 1, 2, ...
          // 7: BSTR WktLinestring = attribute Points3D | example : "LINESTRING(PosX2 PosY2 PosZ2, ..., PosXn-1 PosYn-1 PosZn-1)' with Pos as double; PosZ is optional; Pos1 and Posn automatically created; No additional points : "LINESTRING EMPTY'
          for (vector<ILinkPtr>::size_type cnt_Enter = 0; cnt_Enter < LinksEnter.size(); ++cnt_Enter) {
            for (vector<ILinkPtr>::size_type cnt_Exit = 0; cnt_Exit < LinksExit.size(); ++cnt_Exit) {
              if (cnt_Enter != cnt_Exit) {
                pVissim->GetNet()->GetLinks()->AddConnector(0
                  , LinksEnter[cnt_Enter]->GetLanes()->GetItemByKey(1), 275
                  , LinksExit[cnt_Exit]->GetLanes()->GetItemByKey(1), 0
                  , 1, L"LINESTRING EMPTY"); // unsigned int Key, ILane* FromLane, double FromPos, ILane* ToLane, double ToPos, unsigned int NumberOfLanes, BSTR WktLinestring
              }
            }
          }

          // Iterate the links (just for fun)
          auto spLinks = pVissim->GetNet()->GetLinks();
          for (auto linkIter = spLinks->GetIterator(); linkIter->GetValid(); linkIter->Next()) {
            ILinkPtr link = linkIter->GetItem();
            unsigned int no = link->GetAttValue(L"No");
          }

          // alternative Route :
          ILinkPtr LinksAlternative = pVissim->GetNet()->GetLinks()->AddLink(0
            , L"LINESTRING(-150.000 -9.000 0.000,-113.009 -15.641 0.000,-75.582 -20.384 0.000,-37.864 -23.231 0.000,-0.000 -24.179 0.000,37.864 -23.231 0.000,75.582 -20.384 0.000,113.009 -15.641 0.000,150.000 -9.000 0.000)", makeVariantSafeArray(3.5)); // unsigned int Key, BSTR WktLinestring, SAFEARRAY(double) LaneWidths
          ILinkPtr ConnToAlternative = pVissim->GetNet()->GetLinks()->AddConnector(0
            , LinksEnter[0]->GetLanes()->GetItemByKey(1), 130
            , LinksAlternative->GetLanes()->GetItemByKey(1).GetInterfacePtr(), 0
            , 1, L"LINESTRING(-166.600 -4.181 0.000,-163.323 -5.284 0.000,-160.114 -6.775 0.000,-156.913 -8.368 0.000,-153.664 -9.779 0.000)"); // unsigned int Key, ILane* FromLane, double FromPos, ILane* ToLane, double ToPos, unsigned int NumberOfLanes, BSTR WktLinestring
          ILinkPtr ConnFromAlternative = pVissim->GetNet()->GetLinks()->AddConnector(0
            , LinksAlternative->GetLanes()->GetItemByKey(1).GetInterfacePtr(), LinksAlternative->GetAttValue(L"Length2D")
            , LinksExit[2]->GetLanes()->GetItemByKey(1), 145
            , 1, L"LINESTRING(153.741 -9.595 0.000,156.974 -7.762 0.000,160.112 -5.631 0.000,163.259 -3.606 0.000,166.520 -2.094 0.000)"); // unsigned int Key, ILane* FromLane, double FromPos, ILane* ToLane, double ToPos, unsigned int NumberOfLanes, BSTR WktLinestring
          LinksAlternative->GetLanes()->AddLane(0, 3.5); // unsigned int Key, double Width

          //--------------------------------------
          // Vehicle Inputs
          //--------------------------------------
          for (vector<ILinkPtr>::size_type cnt_Enter = 0; cnt_Enter < LinksEnter.size(); ++cnt_Enter) {
            IVehicleInputPtr VehInput = pVissim->GetNet()->GetVehicleInputs()->AddVehicleInput(0, LinksEnter[cnt_Enter]); // unsigned int Key, Link,
            VehInput->PutAttValue(L"Volume(1)", 500);
          }

          //--------------------------------------
          // Parking Lot
          //--------------------------------------
          IParkingLotPtr ParkingLot1 = pVissim->GetNet()->GetParkingLots()->AddParkingLot(1, LinksExit[2]->GetLanes()->GetItemByKey(1).GetInterfacePtr(), 186); // unsigned int Key, ILane* Lane, double Pos

          IZonePtr zone = pVissim->GetNet()->GetZones()->AddZone(0); // unsigned int Key
          pVissim->GetNet()->GetParkingLots()->AddAbstractParkingLot(0, LinksExit[3], 186, pVissim->GetNet()->GetZones()->GetItemByKey(1).GetInterfacePtr()); // unsigned int Key, ILink* Link, double Pos, IZone* Zone

          //--------------------------------------
          // Matrices
          //--------------------------------------
          pVissim->GetNet()->GetMatrices()->AddMatrix(0); // unsigned int Key

          //--------------------------------------
          // Vehicle Routing Decisions
          //--------------------------------------
          IVehicleRoutingDecisionStaticPtr VehRoutDesSta = pVissim->GetNet()->GetVehicleRoutingDecisionsStatic()->AddVehicleRoutingDecisionStatic(0, LinksEnter[0], 100); // unsigned int Key, ILink* Link, , double Pos
          IVehicleRoutingDecisionPartialPtr VehRoutDesPar = pVissim->GetNet()->GetVehicleRoutingDecisionsPartial()->AddVehicleRoutingDecisionPartial(0, LinksEnter[0], 102, LinksExit[2], 202); // unsigned int Key, ILink* Link, double Pos, ILink* DestLink, double DestPos
          IVehicleRoutingDecisionPartialPTPtr VehRoutDesPPT = pVissim->GetNet()->GetVehicleRoutingDecisionsPartialPT()->AddVehicleRoutingDecisionPartialPT(0, LinksEnter[0], 104, LinksExit[2], 204); // unsigned int Key, ILink* Link, double Pos, ILink* DestLink, double DestPos
          IVehicleRoutingDecisionParkingPtr VehRoutDesPark = pVissim->GetNet()->GetVehicleRoutingDecisionsParking()->AddVehicleRoutingDecisionParking(0, LinksEnter[0], 106); // unsigned int Key, ILink* Link, double Pos
          IVehicleRoutingDecisionManagedLanesPtr VehRoutDesML = pVissim->GetNet()->GetVehicleRoutingDecisionsManagedLanes()->AddVehicleRoutingDecisionManagedLanes(0, LinksEnter[0], 108, LinksExit[2], 208); // unsigned int Key, ILink* Link, double Pos, ILink* DestLink, double DestPos
          IVehicleRoutingDecisionDynamicPtr VehRoutDesDyn = pVissim->GetNet()->GetVehicleRoutingDecisionsDynamic()->AddVehicleRoutingDecisionDynamic(0, LinksEnter[0], 110); // unsigned int Key, ILink* Link, double Pos
          IVehicleRoutingDecisionClosurePtr VehRoutDesClose = pVissim->GetNet()->GetVehicleRoutingDecisionsClosure()->AddVehicleRoutingDecisionClosure(0, LinksEnter[0], 112); // unsigned int Key, ILink* Link, double Pos

          //--------------------------------------
          // Vehicle Routes
          //--------------------------------------
          IVehicleRouteStaticPtr VehRoutSta1 = VehRoutDesSta->GetVehRoutSta()->AddVehicleRouteStatic(0, LinksExit[2], 200); // unsigned int Key, ILink* DestLink, double DestPos
          IVehicleRouteStaticPtr VehRoutSta2 = VehRoutDesSta->GetVehRoutSta()->AddVehicleRouteStatic(0, LinksExit[2], 200); // unsigned int Key, ILink* DestLink, double DestPos
          VehRoutSta2->UpdateLinkSequence(makeVariantSafeArray(LinksAlternative.GetInterfacePtr(), 20)); // SAFEARRAY(VARIANT) RouteSegments Recomputes the link sequence.RouteSegments must be a list of each link and its offset that should be part of the route.
          VehRoutSta2->UpdateLinkSequence(makeVariantSafeArray(ConnToAlternative.GetInterfacePtr(), 1, LinksAlternative.GetInterfacePtr(), 20, ConnFromAlternative.GetInterfacePtr(), 1)); // SAFEARRAY(VARIANT) RouteSegments

          IVehicleRoutePartialPtr VehRoutPar = VehRoutDesPar->GetVehRoutPart()->AddVehicleRoutePartial(0); // unsigned int Key
          VehRoutDesPPT->GetVehRoutPartPT()->AddVehicleRoutePartialPT(0); // unsigned int Key
          VehRoutDesPark->GetVehRoutPark()->AddVehicleRouteParking(0, ParkingLot1); // unsigned int Key, IParkingLot* ParkLot
          VehRoutDesML->GetVehRoutMngLns()->AddVehicleRouteManagedLanes(0, ManagedLaneType::ManagedLaneTypeManaged); // unsigned int Key, enum ManagedLaneType Type[1:Managed; 2: General purpose]
          VehRoutDesClose->GetVehRoutClos()->AddVehicleRouteClosure(0, LinksExit[2], 204); // unsigned int Key, ILink* DestLink, double DestPos

          ILanePtr Lane1 = LinksEnter[0]->GetLanes()->GetItemByKey(1);
          ILanePtr Lane2 = LinksEnter[1]->GetLanes()->GetItemByKey(1);
          ILanePtr Lane3 = LinksEnter[2]->GetLanes()->GetItemByKey(1);
          ILanePtr Lane4 = LinksEnter[3]->GetLanes()->GetItemByKey(1);

          //--------------------------------------
          // Stop Sign
          //--------------------------------------
          pVissim->GetNet()->GetStopSigns()->AddStopSign(0, Lane1, 274); // unsigned int Key, ILane* Lane, double Pos

          //--------------------------------------
          // Priority Rule
          //--------------------------------------
          IPriorityRulePtr PriRule = pVissim->GetNet()->GetPriorityRules()->AddPriorityRule(0, Lane1, 270, makeVariantSafeArray(Lane2.GetInterfacePtr(), 270, Lane3.GetInterfacePtr(), 270)); // unsigned int Key, ILane* Lane, double Pos, SAFEARRAY(VARIANT) ConflictMarkers(ConflictMarkers must be a list that specifies for each ConflictMarker alternating its lane and position)
          PriRule->GetConflictMarkers()->AddConflictMarker(0, Lane4, 270); // unsigned int Key, ILane* Lane, double Pos

          //--------------------------------------
          // Desired Speed Decision
          //--------------------------------------
          pVissim->GetNet()->GetDesSpeedDecisions()->AddDesSpeedDecision(0, Lane1, 10); // int Key, ILane* Lane, double Pos

          //--------------------------------------
          // Reduced Speed Area
          //--------------------------------------
          pVissim->GetNet()->GetReducedSpeedAreas()->AddReducedSpeedArea(0, Lane1, 10); // unsigned int Key, ILane* Lane, double Pos

          //--------------------------------------
          // Node
          //--------------------------------------
          pVissim->GetNet()->GetNodes()->AddNode(0, L"POLYGON((-30 -30, 30 -30, 30 30, -30 30, -30 -30))"); // unsigned int Key, BSTR WktPolygon

          //--------------------------------------
          // Public Transport
          //--------------------------------------
          pVissim->GetNet()->GetPTStops()->AddPTStop(0, Lane1, 150); // unsigned int Key, ILane* Lane, double Pos
          IPTLinePtr PTLine = pVissim->GetNet()->GetPTLines()->AddPTLine(0, LinksEnter[0], LinksExit[2], 200); // unsigned int Key, ILink* EntryLink, ILink* DestLink, double DestPos
          PTLine->GetDepTimes()->AddDepartureTime(0, 123); // unsigned int Key, double Dep
          PTLine->GetDepTimes()->AddDepartureTime(0, 246); // unsigned int Key, double Dep
          PTLine->UpdateLinkSequence(makeVariantSafeArray(LinksAlternative.GetInterfacePtr(), 20)); // SAFEARRAY(VARIANT) RouteSegments Recomputes the link sequence.RouteSegments must be a list of each link and its offset that should be part of the route.
          // PTLine.UpdateLinkSequence([ConnToAlternative, 1, LinksAlternative, 20, ConnFromAlternative, 1]) // SAFEARRAY(VARIANT)

          //--------------------------------------
          // Static 3D Models
          //--------------------------------------
          wstring Static3DModelFile = PTVVissimInstallationPath + L"\\3DModels\\Vegetation\\Tree 04.v3d";
          pVissim->GetNet()->GetStatic3DModels()->AddStatic3DModel(0, Static3DModelFile.c_str(), L"Point(-100, -100, 0)");// unsigned int Key, BSTR ModelFilename, BSTR PosWktPoint3D

          //--------------------------------------
          // Background Image
          //--------------------------------------
          wstring BackgroundImageFile = PTVVissimInstallationPath + L"\\3DModels\\Textures\\Marsh.png";
          pVissim->GetNet()->GetBackgroundImages()->AddBackgroundImage(0, BackgroundImageFile.c_str(), L"Point(100, 100)", L"Point(-100, -100)"); // unsigned int Key, BSTR PathFilename, BSTR PosTRWktPoint, BSTR PosBLWktPoint

          //--------------------------------------
          // Presentation
          //--------------------------------------
          ICameraPositionPtr CamPos = pVissim->GetNet()->GetCameraPositions()->AddCameraPosition(0, L"Point(-100, -100, 100)"); // unsigned int Key, BSTR WktPoint3D
          CamPos->PutAttValue(L"Name", L"Position 1 (added by COM)");
          IStoryboardPtr Storyboard = pVissim->GetNet()->GetStoryboards()->AddStoryboard(0); // unsigned int Key
          Storyboard->GetKeyframes()->AddKeyframe(0); // unsigned int Key

          //--------------------------------------
          // Section
          //--------------------------------------
          ISectionPtr section = pVissim->GetNet()->GetSections()->AddSection(0, L"POLYGON((-200 -150,  200 -150,  200 150,  -200 150))"); // unsigned int Key, BSTR WktPolygon

          //--------------------------------------
          // Signal Controller
          //--------------------------------------
          ISignalControllerPtr SignalController = pVissim->GetNet()->GetSignalControllers()->AddSignalController(0); // unsigned int Key
          SignalController->GetSGs()->AddSignalGroup(0); // unsigned int Key

          //--------------------------------------
          // Detector
          //--------------------------------------
          pVissim->GetNet()->GetDetectors()->AddDetector(0, Lane1, 250); // unsigned int Key, Lane, double Pos

          //--------------------------------------
          // Signal Head
          //--------------------------------------
          pVissim->GetNet()->GetSignalHeads()->AddSignalHead(0, Lane1, 274); // unsigned int Key, ILane* Lane, double Pos

          //--------------------------------------
          // 3D Traffic Signal
          //--------------------------------------
          ITrafficSignal3DPtr TrafficSignal3D = pVissim->GetNet()->GetSignals3D()->AddTrafficSignal3D(0, L"POINT(0 0)"); // unsigned int Key, WktPos3D As String
          TrafficSignal3D->GetSigHeads()->AddSignalHead3D(0); // unsigned int Key
          TrafficSignal3D->GetSigArms()->AddSignalArm3D(0); // unsigned int Key
          TrafficSignal3D->GetTrafficSigns()->AddTrafficSign3D(0); // unsigned int Key
          TrafficSignal3D->GetStreetlights()->AddStreetlight3D(0); // unsigned int Key

          //--------------------------------------
          // Pavement Marking
          //--------------------------------------
          pVissim->GetNet()->GetPavementMarkings()->AddPavementMarking(0, Lane1, 215); // unsigned int Key, ILane* Lane, double Pos

          //--------------------------------------
          // Vehicles Evaluations
          //--------------------------------------
          pVissim->GetNet()->GetDataCollectionPoints()->AddDataCollectionPoint(0, Lane1, 230); // unsigned int Key, ILane* Lane, double Pos
          pVissim->GetNet()->GetVehicleTravelTimeMeasurements()->AddVehicleTravelTimeMeasurement(0, LinksEnter[1], 50, LinksExit[0], 200); // unsigned int Key, ILink* StartLink, double StartPos, ILink* EndLink, double EndPos
          pVissim->GetNet()->GetQueueCounters()->AddQueueCounter(0, LinksEnter[0], 123); // unsigned int Key, ILink* Link, double Pos
          pVissim->GetNet()->GetDelayMeasurements()->AddDelayMeasurement(0); // unsigned int Key




          //=============================================
          //=============================================
          // ADDING OBJECTS SPECIFICALLY FOR PEDESTRIANS
          //=============================================
          //=============================================

          //--------------------------------------
          // Pedestrian Base Data
          //--------------------------------------
          pVissim->GetNet()->GetPedestrianTypes()->AddPedestrianType(0); // unsigned int Key
          pVissim->GetNet()->GetPedestrianClasses()->AddPedestrianClass(0); // unsigned int Key
          pVissim->GetNet()->GetWalkingBehaviors()->AddWalkingBehavior(0); // unsigned int Key
          IAreaBehaviorTypePtr AreaBehaviorType = pVissim->GetNet()->GetAreaBehaviorTypes()->AddAreaBehaviorType(0); // unsigned int Key
          AreaBehaviorType->GetAreaBehavTypeElements()->AddAreaBehaviorTypeElement(pVissim->GetNet()->GetPedestrianClasses()->GetItemByKey(30).GetInterfacePtr()); // IPedestrianClass* PedClass

          //--------------------------------------
          // Pedestrian Compositions
          //--------------------------------------
          pVissim->GetNet()->GetPedestrianCompositions()->AddPedestrianComposition(0
            , makeVariantSafeArray()); // unsigned int Key, SAFEARRAY(VARIANT) PedCompRelFlows
          pVissim->GetNet()->GetPedestrianCompositions()->AddPedestrianComposition(9
            , makeVariantSafeArray(pVissim->GetNet()->GetPedestrianTypes()->GetItemByKey(100).GetInterfacePtr(), pVissim->GetNet()->GetDesSpeedDistributions()->GetItemByKey(5).GetInterfacePtr(), pVissim->GetNet()->GetPedestrianTypes()->GetItemByKey(200).GetInterfacePtr(), pVissim->GetNet()->GetDesSpeedDistributions()->GetItemByKey(5).GetInterfacePtr())); // unsigned int Key, SAFEARRAY(VARIANT) PedCompRelFlows
          pVissim->GetNet()->GetPedestrianCompositions()->GetItemByKey(9)->GetPedCompRelFlows()->AddPedestrianCompositionRelativeFlow(pVissim->GetNet()->GetPedestrianTypes()->GetItemByKey(300).GetInterfacePtr(), pVissim->GetNet()->GetDesSpeedDistributions()->GetItemByKey(5).GetInterfacePtr()); // IPedestrianType* PedType, IDesSpeedDistribution* DesSpeedDistr

          //--------------------------------------
          // Areas
          //--------------------------------------
          // Input parameters of AddArea :
          // 1: unsigned int Key = attribute Number(No) | example : 0 or 123
          // 2: BSTR WktPolygon = attribute Points | example : "POLYGON(PosX1 PosY1, PosX2 PosY2, ..., PosXn PosYn)' with Pos as double

          // Example : Area1 = pVissim->GetNet()->GetAreas()->AddArea(0, "POLYGON((0 50, 50 50, 50 100, 0 100))') // unsigned int Key, BSTR WktPolygon

          // Create some areas :
          vector<double> xydim{ 10, 10 };
          vector<vector<double>> StartPosAll{ { 50, 50 }, { 50, 60 }, { 50, 70 }, { 50, 80 }, { 50, 90 }, { 60, 60 }, { 70, 60 }, { 70, 70 }, { 70, 80 }, { 60, 80 } };
          vector<IAreaPtr> Area(StartPosAll.size() + 1);
          vector<IAreaPtr>::size_type cnt = 0;
          for (auto StartPos : StartPosAll) {
            double x1 = StartPos[0];
            double x2 = StartPos[0] + xydim[0];
            double y1 = StartPos[1];
            double y2 = StartPos[1] + xydim[1];
            wostringstream stringStream;
            stringStream << L"POLYGON((" << x1 << " " << y1 << ", " << x1 << " " << y2 << ", " << x2 << " " << y2 << ", " << x2 << " " << y1 << "))";
            wstring Polyg = stringStream.str();
            Area[cnt] = pVissim->GetNet()->GetAreas()->AddArea(0, Polyg.c_str()); // unsigned int Key, BSTR WktPolygon
            cnt = cnt + 1;
          }

          //--------------------------------------
          // Obstacle & Ramp
          //--------------------------------------
          pVissim->GetNet()->GetObstacles()->AddObstacle(0, L"POLYGON((55 72.5, 65 72.5, 65 77.5, 55 77.5))"); // unsigned int Key, BSTR WktPolygon
          IRampPtr Ramp = pVissim->GetNet()->GetRamps()->AddRamp(0, L"POLYGON((50 100,  50 110,  60 110,  60 100))"); // unsigned int Key, BSTR WktPolygon
          Area[cnt] = pVissim->GetNet()->GetAreas()->AddArea(0, L"POLYGON((50 110,  60 110,  60 120,  50 120))"); // unsigned int Key, BSTR WktPolygon

          //--------------------------------------
          // Pedestrian Input
          //--------------------------------------
          pVissim->GetNet()->GetPedestrianInputs()->AddPedestrianInput(0, Area[0]); // unsigned int Key, IArea* Area

          //--------------------------------------
          // Pedestrian Routing Decisions & Routes
          //--------------------------------------
          IPedestrianRoutingDecisionStaticPtr PedRouteDesSta1 = pVissim->GetNet()->GetPedestrianRoutingDecisionsStatic()->AddPedestrianRoutingDecisionStatic(0, Area[0]); // unsigned int Key, IArea* Area
          IPedestrianRouteStaticPtr PedRouteSta1 = PedRouteDesSta1->GetPedRoutSta()->AddPedestrianRouteStaticOnArea(0, Area[1]); // unsigned int Key, IArea* Area
          PedRouteSta1->GetPedRoutLoc()->AddPedestrianRouteLocationOnArea(0, Area[2]); // unsigned int Key, IArea* Area | = > will become next route end
          PedRouteSta1->GetPedRoutLoc()->AddPedestrianRouteLocationOnArea(0, Area[5]); // unsigned int Key, IArea* Area | = > will become next route end
          PedRouteSta1->GetPedRoutLoc()->AddPedestrianRouteLocationOnArea(0, Area[3]); // unsigned int Key, IArea* Area | = > will become next route end
          PedRouteSta1->GetPedRoutLoc()->AddPedestrianRouteLocationOnArea(0, Area[4]); // unsigned int Key, IArea* Area | = > will become next route end

          IPedestrianRouteStaticPtr PedRouteSta2 = PedRouteDesSta1->GetPedRoutSta()->AddPedestrianRouteStaticOnRamp(0, Ramp);  // unsigned int Key, IRamp* PedRoutLoc
          IPedestrianRouteStaticPtr PedRouteSta3 = PedRouteDesSta1->GetPedRoutSta()->AddPedestrianRouteStaticOnArea(0, Area[9]); // unsigned int Key, IArea* Area
          PedRouteSta3->GetPedRoutLoc()->AddPedestrianRouteLocationOnRamp(0, Ramp); // unsigned int Key, IRamp* PedRoutLoc | = > will become next route end

          IPedestrianRoutingDecisionPartialPtr PedRouteDesPart1 = pVissim->GetNet()->GetPedestrianRoutingDecisionsPartial()->AddPedestrianRoutingDecisionPartial(0, Area[1]); // unsigned int Key, IArea* Area
          IPedestrianRoutePartialPtr PedRoutePart1 = PedRouteDesPart1->GetPedRoutPart()->AddPedestrianRoutePartialOnArea(0, Area[7]); // unsigned int Key, IArea* Area
          PedRoutePart1->GetPedRoutLoc()->AddPedestrianRouteLocationOnArea(0, Area[3]); // unsigned int Key, IArea* Area | = > will become next route end
          IPedestrianRoutePartialPtr PedRoutePart2 = PedRouteDesPart1->GetPedRoutPart()->AddPedestrianRoutePartialOnArea(0, Area[3]); // unsigned int Key, IArea* Area

          IPedestrianRoutingDecisionPartialPtr PedRouteDesPart2 = pVissim->GetNet()->GetPedestrianRoutingDecisionsPartial()->AddPedestrianRoutingDecisionPartial(0, Area[1]); // unsigned int Key, IArea* Area
          PedRoutePart2 = PedRouteDesPart2->GetPedRoutPart()->AddPedestrianRoutePartialOnRamp(0, Ramp);  // unsigned int Key, IRamp* PedRoutLoc

          //--------------------------------------
          // Pedestrian Evaluation
          //--------------------------------------
          pVissim->GetNet()->GetPedestrianTravelTimeMeasurements()->AddPedestrianTravelTimeMeasurement(0, Area[0], Area[4]); // unsigned int Key, IArea* StartArea, IArea* EndArea


          // free Vissim object
          if (pVissim.GetInterfacePtr() != NULL) {
            pVissim.Detach()->Release();
          }

        }
        catch (_com_error & e) {
          cout << "COM error: " << e.ErrorMessage() << endl;
          return -1;
        }
      }
      // uninitialize COM
      CoUninitialize();
      return 0;
    }
	}
	else {
		// TODO: change error code to suit your needs
		_tprintf(_T("Fatal Error: GetModuleHandle failed\n"));
		nRetCode = 1;
	}

	return nRetCode;
}
