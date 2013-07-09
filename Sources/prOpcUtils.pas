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
unit prOpcUtils;
{$I prOpcCompilerDirectives.inc}

{History
  13-03-01. Dropped PathChar property. I can see real problems with this,
            and I think it is very unlikely to be ever needed.
  14-03-10. Numerous mods to handle new exception scheme.


Release 1.14
1.14.1      Fixed a few memory leaks caused by passing structures by value (!).
            I lecture everyone endlessly about this and then get caught out
            myself! This bug may be an issue at RR Derby.
}

{do not remove the following comment}

//CE_Desc_Include(..\Help\Topics.txt)

interface
uses
  SysUtils, Windows, ActiveX, Classes, prOpcDa, prOpcComn, prOpcTypes,
  prOpcError;

type
  DAVariant = array of OleVariant;

procedure FreeOPCItemState(var v: OPCITEMSTATE);
procedure FreeOPCItemStateArray(dwCount: UINT; pv: POPCITEMSTATEARRAY);

procedure FreeOPCItemResult(var v: OPCITEMRESULT);
procedure FreeOPCItemResultArray(dwCount: UINT; pv: POPCITEMRESULTARRAY);

procedure FreeOPCItemAttributes(var v: OPCITEMATTRIBUTES);
procedure FreeOPCItemAttributesArray(dwCount: UINT; pv: POPCITEMATTRIBUTESARRAY);

procedure FreeOPCItemProperties(var v: OPCITEMPROPERTIES);

procedure FreeOPCItemBrowseElement(v: OPCBROWSEELEMENT);
procedure FreeOPCItemBrowseElementArray(dwCount: UINT; pv: POPCBROWSEELEMENTARRAY);

procedure FreeOleVariantArray(dwCount: Integer; pv: POleVariantArray);

procedure FreeResultList(pv: PResultList);

procedure FreeServerStatus(pv: POPCSERVERSTATUS);

procedure FormatOPCStmDataTime(Stream: Pointer; var GroupHeader: POPCGROUPHEADER;
                                var ItemHeaders: POPCITEMHEADER1ARRAY; var Data: DAVariant);
procedure FormatOPCStmData(Stream: Pointer; var GroupHeader: POPCGROUPHEADER;
                            var ItemHeaders: POPCITEMHEADER2ARRAY; var Data: DAVariant);
procedure FormatOPCStmWriteComplete(Stream: Pointer; var GroupHeader: POPCGROUPHEADERWRITE;
                                     var ItemHeaders: POPCITEMHEADERWRITEARRAY);

function GetMalloc: IMalloc;

function GIT : IGlobalInterfaceTable;

function DatatypeToStr(Datatype: Integer): string;
function QualityToStr(Quality: Word): String;
function AccessRightsToStr(AccessRights: DWORD): String;
procedure InitialiseClientSecurity;

function GetOpcErrorString(Server: IUnknown; Code: HRESULT): string;

implementation
uses
{$IFDEF D6UP}
  Variants,
{$ENDIF}
  ComObj;

resourcestring
  SUnknownError = 'Unknown error code %.8x';

const
  CLSID_StdGlobalInterfaceTable : TGUID = '{00000323-0000-0000-C000-000000000046}';

var
  cGIT : IGlobalInterfaceTable = nil;

function GIT : IGlobalInterfaceTable;
begin
  if (cGIT = nil) then
    OleCheck(CoCreateInstance(CLSID_StdGlobalInterfaceTable, nil, CLSCTX_ALL,
      IGlobalInterfaceTable, cGIT));
  Result := cGIT;
end;

function GetMalloc: IMalloc;
begin
  OleCheck(CoGetMalloc(1, Result))
end;

procedure FreeOPCItemState( var v: OPCITEMSTATE);
begin
  v.vDataValue:= Unassigned
end;

procedure FreeOPCItemStateArray( dwCount: UINT; pv: POPCITEMSTATEARRAY);
var
  i: Integer;
begin
  if Assigned(pv) then
  begin
    for i:= 0 to dwCount - 1 do
      FreeOPCItemState(pv^[i]);
    CoTaskMemFree(pv)
  end
end;

procedure FreeOPCItemResult( var v: OPCITEMRESULT);
begin
  with v do if Assigned(pBlob) then
    CoTaskMemFree(pBlob)
end;

procedure FreeOPCItemResultArray( dwCount: UINT; pv: POPCITEMRESULTARRAY);
var
  i: Integer;
begin
  if Assigned(pv) then
  begin
    for i:= 0 to dwCount - 1 do
      FreeOPCItemResult(pv^[i]);
    CoTaskMemFree(pv)
  end
end;

procedure FreeOPCItemAttributes( var v: OPCITEMATTRIBUTES);
begin
  with v do
  with GetMalloc do
  begin
    Free(szAccessPath);
    Free(szItemID);
    Free(pBlob);
    vEUInfo:= Unassigned
  end
end;

procedure FreeOPCItemAttributesArray( dwCount: UINT; pv: POPCITEMATTRIBUTESARRAY);
var
  i: Integer;
begin
  if Assigned(pv) then
  begin
    for i:= 0 to dwCount - 1 do
      FreeOPCItemAttributes(pv^[i]);
    CoTaskMemFree(pv)
  end
end;

procedure FreeOPCItemProperties(var v: OPCITEMPROPERTIES);
var
  i: Integer;
begin
  with v do
  with GetMalloc do
  begin
    if dwNumProperties > 0 then
      for i:= 0 to dwNumProperties - 1 do
      begin
        Free(pItemProperties[i].szItemID);
        Free(pItemProperties[i].szDescription);
        pItemProperties[i].vValue := Unassigned;
      end;
    Free(pItemProperties);
  end;
end;

procedure FreeOPCItemBrowseElement(v: OPCBROWSEELEMENT);
begin
  with v do
  with GetMalloc do
  begin
    Free(szName);
    Free(szItemID);
    FreeOPCItemProperties(ItemProperties);
  end
end;

procedure FreeOPCItemBrowseElementArray(dwCount: UINT; pv: POPCBROWSEELEMENTARRAY);
var
  i: Integer;
begin
  if Assigned(pv) then
  begin
    if dwCount > 0 then
      for i:= 0 to dwCount - 1 do
        FreeOPCItemBrowseElement(pv^[i]);
    CoTaskMemFree(pv)
  end
end;

procedure FreeOleVariantArray( dwCount: Integer; pv: POleVariantArray);
var
  i: Integer;
begin
  if Assigned(pv) then
  begin
    for i:= 0 to dwCount - 1 do
      pv^[i]:= Unassigned;
    CoTaskMemFree(pv)
  end
end;

procedure FreeResultList(pv: PResultList);
begin
  if Assigned(pv) then
    CoTaskMemFree(pv)
end;

procedure FreeServerStatus(pv: POPCSERVERSTATUS);
begin
  CoTaskMemFree(pv.szVendorInfo);
  CoTaskMemFree(pv);
end;

type
  TDataPtr = record
  case Byte of
    0: (Data: POleVariant);
    1: (StrLen: PDWord);
    2: (OleStr: PWideChar);
    3: (VarArray: PVarArray);
  end;

procedure ReadVariant(Stream: Integer;  var Data: OleVariant);
var
  DataPtr: TDataPtr absolute Stream;
  StmData: POleVariant;
begin
  StmData:= DataPtr.Data;
  with TVarData(StmData^) do
  case VType of
    varOleStr:
    begin
      Inc(Stream, SizeOf(TVarData) + SizeOf(DWORD));
      VOleStr:= DataPtr.OleStr
    end;
    varArray:    {I have been unable to test this}
    begin
      Inc(Stream, SizeOf(TVarData));
      VArray:= DataPtr.VarArray;
    end
  end;
  Data:= StmData^
end;

procedure FormatOPCStmDataTime( Stream: Pointer; var GroupHeader: POPCGROUPHEADER;
                                var ItemHeaders: POPCITEMHEADER1ARRAY; var Data: DAVariant);
var
  StmInt: DWORD absolute Stream;
  i: Integer;
begin
  GroupHeader:= Stream;
  ItemHeaders:= Pointer(StmInt + SizeOf(OPCGROUPHEADER));
  SetLength(Data, GroupHeader^.dwItemCount);
  for i:= 0 to GroupHeader^.dwItemCount - 1 do
    ReadVariant(StmInt + ItemHeaders^[i].dwValueOffset, Data[i])
end;

procedure FormatOPCStmData( Stream: Pointer; var GroupHeader: POPCGROUPHEADER;
                            var ItemHeaders: POPCITEMHEADER2ARRAY; var Data: DAVariant);
var
  StmInt: DWORD absolute Stream;
  i: Integer;
begin
  GroupHeader:= Stream;
  ItemHeaders:= Pointer(StmInt + SizeOf(OPCGROUPHEADER));
  SetLength(Data, GroupHeader^.dwItemCount);
  for i:= 0 to GroupHeader^.dwItemCount - 1 do
    ReadVariant(StmInt + ItemHeaders^[i].dwValueOffset, Data[i])
end;

procedure FormatOPCStmWriteComplete( Stream: Pointer; var GroupHeader: POPCGROUPHEADERWRITE;
                                     var ItemHeaders: POPCITEMHEADERWRITEARRAY);
var
  StmInt: Integer absolute Stream;
begin
  GroupHeader:= Stream;
  inc(StmInt, SizeOf(OPCGROUPHEADER));
  ItemHeaders:= Stream
end;

function DatatypeToStr(Datatype: Integer): string;
begin
  case Datatype of
    varEmpty: Result:= 'Empty';
    varNull: Result:= 'Null';
    varSmallint: Result:= 'Smallint';
    varInteger: Result:= 'Integer';
    varSingle: Result:= 'Single';
    varDouble: Result:= 'Double';
    varCurrency: Result:= 'Currency';
    varDate: Result:= 'Date';
    varOleStr: Result:= 'OleStr';
    varDispatch: Result:= 'Dispatch';
    varError: Result:= 'Error';
    varBoolean: Result:= 'Boolean';
    varVariant: Result:= 'Variant';
    varShortInt: Result := 'ShortInt';
    varByte: Result := 'Byte';
    varWord: Result := 'Word';
    varLongWord: Result := 'LongWord';
    varInt64: Result := 'Int64';
  else
    Result:= 'Unknown';
  end
end;

function QualityToStr(Quality: Word): String;
begin
  case (Quality and OPC_QUALITY_MASK) of
    OPC_QUALITY_BAD:
    begin
      Result:= 'Bad ';
      case Quality and OPC_STATUS_MASK of
        OPC_QUALITY_CONFIG_ERROR: Result:= Result + 'Config error ';
        OPC_QUALITY_NOT_CONNECTED: Result:= Result + 'not connected ';
        OPC_QUALITY_DEVICE_FAILURE: Result:= Result + 'device failure ';
        OPC_QUALITY_SENSOR_FAILURE: Result:= Result + 'sensor failure ';
        OPC_QUALITY_LAST_KNOWN: Result:= Result + 'last known ';
        OPC_QUALITY_COMM_FAILURE: Result:= Result + 'comm failure ';
        OPC_QUALITY_OUT_OF_SERVICE: Result:= Result + 'out of service ';
      end
    end;
    OPC_QUALITY_UNCERTAIN:
    begin
      Result:= 'Uncertain ';
      case Quality and OPC_STATUS_MASK of
        OPC_QUALITY_LAST_USABLE: Result:= Result + 'last usable ';
        OPC_QUALITY_SENSOR_CAL: Result:= Result + 'sensor cal ';
        OPC_QUALITY_EGU_EXCEEDED: Result:= Result + 'egu exceeded ';
        OPC_QUALITY_SUB_NORMAL: Result:= Result + 'sub normal ';
      end
    end;
    OPC_QUALITY_GOOD:
    begin
      Result:= 'Good ';
      case Quality and OPC_STATUS_MASK of
        OPC_QUALITY_LOCAL_OVERRIDE: Result:= Result + 'local override ';
      end
    end;
  end;
  case Quality and OPC_LIMIT_MASK of
    OPC_LIMIT_OK: Result:= Result + 'limit ok';
    OPC_LIMIT_LOW: Result:= Result + 'limit low';
    OPC_LIMIT_HIGH: Result:= Result + 'limit high';
    OPC_LIMIT_CONST: Result:= Result + 'limit const';
  end
end;

function AccessRightsToStr(AccessRights: DWORD): String;
begin
  Result:= '';
  if (AccessRights and OPC_READABLE) <> 0 then
    Result:= Result + 'r';
  if (AccessRights and OPC_WRITABLE) <> 0 then
    Result:= Result + 'w'
end;


procedure InitialiseClientSecurity;
const
  RPC_C_AUTHN_LEVEL_NONE = 1;
  RPC_C_IMP_LEVEL_IMPERSONATE = 3;
  EOAC_NONE = 0;
var
  Res: HRESULT;
begin
  Res:= CoInitializeSecurity(nil, -1, nil, nil,
    RPC_C_AUTHN_LEVEL_NONE,
    RPC_C_IMP_LEVEL_IMPERSONATE, nil, EOAC_NONE, nil);
  if (Res <> RPC_E_TOO_LATE) then
    OleCheck(Res)
end;

function GetOpcErrorString(Server: IUnknown; Code: HRESULT): string;
{this function is for clients...}
var
  Common: IOPCCommon;
  ppString: PWideChar;
  Buf: array[0..255] of Char;
begin
  if not StdOpcErrorToStr(Code, Result) then
  begin
    if Assigned(Server) and  {ask the server}
       (Server.QueryInterface(IOPCCommon, Common) = S_OK) and
       (Common.GetErrorString(Code, ppString) = S_OK) then
    begin
      Result:= ppString;
      CoTaskMemFree(ppString)
    end else
    if FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM or FORMAT_MESSAGE_ARGUMENT_ARRAY,
          nil, DWORD(Code), 0, Buf, SizeOf(Buf), nil) > 0 then
    begin
      Result:= Buf
    end else
    begin  {give up}
      FmtStr(Result, SUnknownError, [Code])
    end
  end
end;

end.
