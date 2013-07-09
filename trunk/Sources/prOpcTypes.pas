{------------------------------------------------------------}
{The MIT License (MIT)

 prOpc Toolkit
 Copyright (c) 2000, 2001 Production Robots Engineering Ltd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.}
{------------------------------------------------------------}
{------------------------------------------------------------}
{ This unit derived substantially from work by:              }
{ OPC Programmers' Connection                                }
{ http://www.opc.dial.pipex.com/                             }
{ mailto:opc@dial.pipex.com                                  }
{------------------------------------------------------------}
{History
1.10.1      Added function IsSupportedVarType
1.10.2      Modifications to PRttiHandle type to support
            ArrayTypes

Release 1.14 11/04/02
1.14.1      Interfaces for EUInfo and ItemProperties added.

}

unit prOpcTypes;
{$I prOpcCompilerDirectives.inc}
interface
uses
  SysUtils, Classes, Windows, TypInfo, ActiveX;

const
  MaxArraySize      = $2000000;

type
  {OPC types shared across multiple definitions}
{  TOleEnum          = type Integer; }

  OPCHANDLE         = DWORD;
  POPCHANDLE        = ^OPCHANDLE;
  OPCHANDLEARRAY    = array[0..MaxArraySize] of OPCHANDLE;
  POPCHANDLEARRAY   = ^OPCHANDLEARRAY;

  PVarType          = ^TVarType;
  TVarTypeList      = array[0..MaxArraySize] of TVarType;
  PVarTypeList      = ^TVarTypeList;

  POleVariant       = ^OleVariant;
  OleVariantArray   = array[0..MaxArraySize] of OleVariant;
  POleVariantArray  = ^OleVariantArray;

  PLCID             = ^TLCID;

  BOOLARRAY         = array[0..MaxArraySize] of BOOL;
  PBOOLARRAY        = ^BOOLARRAY;

  WORDARRAY         = array[0..MaxArraySize] of Word;
  PWORDARRAY        = ^WORDARRAY;

  DWORDARRAY        = array[0..MaxArraySize] of DWORD;
  PDWORDARRAY       = ^DWORDARRAY;

  SingleArray       = array[0..MaxArraySize] of Single;
  PSingleArray      = ^SingleArray;

  TFileTimeArray    = array[0..MaxArraySize] of TFileTime;
  PFileTimeArray    = ^TFileTimeArray;

  {pascal types}
  TAccessRight = (iaRead, iaWrite);
  TAccessRights = set of TAccessRight;

  TOpcDataAccessType = (opcDA1, opcDA2, opcDA3);
  TOpcDataAccessTypes = set of TOpcDataAccessType;

  TEuType = (euNone, euAnalog, euEnumerated);

  TArraySyntax = (asNone, asComma, asBrackets);

  PArrayProperty = ^TArrayProperty;
  TArrayProperty = record
    Size: Integer;
    _Type: Integer;
    Index: Integer;  {-1 for entire array}
    GetProc, SetProc: TMethod;
  end;

  IEUInfo = interface
    function EUType: TEuType;
    function EUInfo: OleVariant;
  end;

  { Item Properties }
  IItemProperty = interface
  ['{12FB60B3-6B72-4CC8-91E2-9E29B1308BF0}']
    function Description: string;
    function DataType: Integer;
    function GetPropertyValue: OleVariant;
    function Pid: LongWord;
  end;

  IItemProperties = interface
  ['{428505DD-103C-4716-AEBC-C5489ACDE49A}']
    function GetPropertyItem(Index: Integer): IItemProperty;
    function GetProperty(Pid: LongWord): IItemProperty;
    procedure Add(const ItemProperty: IItemProperty);
    function Count: Integer;
  end;

function OpcAccessRights(AccessRights: TAccessRights): DWORD;
function NativeAccessRights(OpcRights: DWORD): TAccessRights;
function FileTimeToDateTime( const Ft: TFiletime): TDateTime;
function DateTimeToFileTime( const Dt: TDateTime): TFileTime;
function IsSupportedVarType( Vtype: Integer): Boolean;

implementation
uses
  ComObj, prOpcError;

function OpcAccessRights(AccessRights: TAccessRights): DWORD;
{OPC_READABLE = 1 and OPC_WRITABLE = 2 so
  AccessRights are binary compatible}
begin
  Result:= 0;
  Move(AccessRights, Result, 1)
end;

function NativeAccessRights(OpcRights: DWORD): TAccessRights;
begin
  Result:= [];
  Move(OpcRights, Result, 1)
end;

function FileTimeToDateTime( const Ft: TFiletime): TDateTime;
var
  St: TSystemTime;
begin
  FiletimeToSystemTime(Ft, St);
  Result:= SystemTimeToDateTime(St)
end;

function DateTimeToFileTime( const Dt: TDateTime): TFileTime;
var
  St: TSystemTime;
begin
  DateTimeToSystemTime(Dt, St);
  SystemTimeToFileTime(St, Result);
end;

function IsSupportedVarType( Vtype: Integer): Boolean;
begin
  case VType and not varArray of
    VT_EMPTY,
    VT_I2,
    VT_I4,
    VT_R4,
    VT_R8,
    VT_CY,
    VT_DATE,
    VT_BSTR,
    VT_BOOL,
    VT_I1,
    VT_UI1,
    VT_UI2,
    VT_UI4,
    VT_INT,
    VT_UINT: Result:= true;
  else
    Result:= false
  end
end;


end.


