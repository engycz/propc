{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
{ This unit derived substantially from work by:              }
{ OPC Programmers' Connection                                }
{ http://www.opcconnect.com/                                 }
{ mailto:opc@dial.pipex.com                                  }
{------------------------------------------------------------}
unit prOpcComn;
{$I prOpcCompilerDirectives.inc}
// ************************************************************************ //
// Type Lib: opccomn_ps.dll
// IID\LCID: {B28EEDB1-AC6F-11D1-84D5-00608CB8A7E9}\0
// ************************************************************************ //
interface

uses
  Windows, ActiveX;

// *********************************************************************//
// GUIDS declared in the TypeLibrary                                    //
// *********************************************************************//
const
  LIBID_OPCCOMN: TGUID = '{B28EEDB1-AC6F-11D1-84D5-00608CB8A7E9}';
  IID_IOPCCommon: TIID = '{F31DFDE2-07B6-11D2-B2D8-0060083BA1FB}';
  IID_IOPCShutdown: TIID = '{F31DFDE1-07B6-11D2-B2D8-0060083BA1FB}';
  IID_IOPCServerList: TIID = '{13486D50-4821-11D2-A494-3CB306C10000}';

  CLSID_OPCServerList: TGUID = '{13486D51-4821-11D2-A494-3CB306C10000}';

type

// *********************************************************************//
// Forward declaration of interfaces defined in Type Library            //
// *********************************************************************//
  IOPCCommon = interface;
  IOPCShutdown = interface;
  IOPCServerList = interface;

// *********************************************************************//
// Declaration of structures, unions and aliases.                       //
// *********************************************************************//
  LCIDARRAY = array[0..65535] of LCID;
  PLCIDARRAY = ^LCIDARRAY;

// *********************************************************************//
// Interface: IOPCCommon
// GUID:      {F31DFDE2-07B6-11D2-B2D8-0060083BA1FB}
// *********************************************************************//
  IOPCCommon = interface(IUnknown)
    ['{F31DFDE2-07B6-11D2-B2D8-0060083BA1FB}']
    function SetLocaleID(
            dwLcid:                     TLCID): HResult; stdcall;
    function GetLocaleID(
      out   pdwLcid:                    TLCID): HResult; stdcall;
    function QueryAvailableLocaleIDs(
      out   pdwCount:                   UINT;
      out   pdwLcid:                    PLCIDARRAY): HResult; stdcall;
    function GetErrorString(
            dwError:                    HResult;
      out   ppString:                   POleStr): HResult; stdcall;
    function SetClientName(
            szName:                     POleStr): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IOPCShutdown
// GUID:      {F31DFDE1-07B6-11D2-B2D8-0060083BA1FB}
// *********************************************************************//
  IOPCShutdown = interface(IUnknown)
    ['{F31DFDE1-07B6-11D2-B2D8-0060083BA1FB}']
    function ShutdownRequest(
            szReason:                   POleStr): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IOPCServerList
// GUID:      {13486D50-4821-11D2-A494-3CB306C10000}
// *********************************************************************//
  IOPCServerList = interface(IUnknown)
    ['{13486D50-4821-11D2-A494-3CB306C10000}']
    function EnumClassesOfCategories(
            cImplemented:               ULONG;
            rgcatidImpl:                PGUID;
            cRequired:                  ULONG;
            rgcatidReq:                 PGUID;
      out   ppenumClsid:                IEnumGUID): HResult; stdcall;
    function GetClassDetails(
      const clsid:                      TCLSID;
      out   ppszProgID:                 POleStr;
      out   ppszUserType:               POleStr): HResult; stdcall;
    function CLSIDFromProgID(
            szProgId:                   POleStr;
      out   clsid:                      TCLSID): HResult; stdcall;
  end;

implementation

end.
