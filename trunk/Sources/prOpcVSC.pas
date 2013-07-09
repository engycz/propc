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
unit prOpcVSC;
{$I prOpcCompilerDirectives.inc}

{History:
File Created for evaluation/test 28/04/02

Release 1.14 01/06/02
  Extensively modified for release

Release 1.15 18/09/02
  cf 1.15.1 Fixed broken D5 support.

}
interface
uses
  Windows, SysUtils, Classes;

type
  TItemAccessRight = (arRead, arWrite);       {clone of prOpcTypes.TAccessRight}
  TItemAccessRights = set of TItemAccessRight;            {"}

{$IFDEF VER130}    {cf 1.15.1}
  TVarType = Word;
{$ENDIF}

  IOpcClient = interface(IDispatch)
    ['{164EF3A1-5525-40EA-BD94-B470D6B21CDA}']
    function GetItemValue(const Name: string): OleVariant;
    procedure SetItemValue(const Name: string; const Value: OleVariant);
    function GetUpdateRate: Integer;
    function GetPercentDeadband: Single;
    procedure SetPercentDeadband(const Value: Single);
    function  GetAccessRights(const Name: string): TItemAccessRights;
    function GetItemTimestamp(const Name: string): TDateTime;
    function GetItemQuality(const Name: string): DWORD;
    function GetItemType(const Name: string): TVarType;

    {Public Interface - Group}
    property UpdateRate: Integer read GetUpdateRate;
    property PercentDeadband: Single read GetPercentDeadband write SetPercentDeadband;

    {Public Interface - Item}
    property ItemAccessRights[const Name: string]: TItemAccessRights read GetAccessRights;
    property ItemType[const Name: string]: TVarType read GetItemType;
    property ItemTimestamp[const Name: string]: TDateTime read GetItemTimestamp;
    property ItemValue[const Name: string]: OleVariant read GetItemValue write SetItemValue; default;
  end;

  {this is raised if an attempt is made to read a tag with bad quality}
  EOpcQuality = class(Exception)
    Quality: DWORD; {see prOpcDa & Spec for details}
    constructor Create(aQuality: DWORD);
  end;

  {use EOleSysError as general exception type}

function OpcClient(
  const HostName, ProgId: string;
  UpdateRate: Integer = 0): IOpcClient;

implementation
uses
{$IFDEF D6}
  Variants,
{$ENDIF}
  ActiveX, prOpcDa, prOpcTypes, prOpcError, prOpcEnum, prOpcComn, ComObj,
  prOpcUtils;

resourcestring
  SNoProgId = 'ProgId must be specified';
  SBadQuality = 'Bad OPC quality $%.8x';

type
  TVSC = class;

  TItemRec = class
    hServer: OPCHANDLE;
    Value: OleVariant;
    Timestamp: TDateTime;  {not used}
    Quality: DWORD;  {not used}
    Owner: TVSC;
    CanonicalDataType: Word;
    AccessRights: TItemAccessRights;
    constructor Create(aOwner: TVSC; const ItemName: string);
    procedure Put(const NewValue: OleVariant);
    function Get: OleVariant;
  end;

  TServer = class
    Server: IOPCServer;
    GroupCount: Integer; {could get this from the server via GroupEnumerator etc}
    ShutdownCookie: Integer;
    ServerDesc, VendorName: string;
    procedure RemoveGroup(hGroup: OPCHANDLE);
    function AddGroup(UpdateRate, ClientHandle: Integer; var ServerHandle: OPCHANDLE): IUnknown;
    destructor Destroy; override;
    class function CreateServer(const aHostName, aProgId: string): TServer;
  end;

{
  TShutdownSink = class(TInterfacedObject, IOPCShutdown)
    Server: TServer;
    function ShutdownRequest(szReason: POleStr): HResult; stdcall;
    constructor Create(aServer: TServer);
  end;
}

  TVSC = class(TInterfacedObject, IOpcClient) {actually a group}
    Group: IUnknown;
    hGroup: OPCHANDLE;
    UpdateRate: Cardinal;
    PercentDeadband: Single;
    Items: TStringList;
    Server: TServer;
    DA2Cookie: Integer;
    DA1Cookie: Integer;
    function GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; stdcall;
    function GetDispId(Name: PWideChar): TItemRec;
    function GetItemRec(const Name: string): TItemRec;
    procedure DataChange(dwCount: DWORD; phClientItems: POPCHANDLEARRAY;
                pvValues: POleVariantArray; pwQualities: PWORDARRAY;
                pftTimeStamps: PFileTimeArray; pErrors: PResultList);
    procedure DA1DataTimeChange(dwCount: DWORD; phClientHeaders: POPCITEMHEADER1ARRAY;
                Values: DAVariant);

    function GetItemValue(const Name: string): OleVariant;
    procedure SetItemValue(const Name: string; const Value: OleVariant);
    function GetUpdateRate: Integer;
    function GetPercentDeadband: Single;
    procedure SetPercentDeadband(const Value: Single);
    function  GetAccessRights(const Name: string): TItemAccessRights;
    function GetItemType(const Name: string): TVarType;
    function GetItemTimestamp(const Name: string): TDateTime;
    function GetItemQuality(const Name: string): DWORD;

    constructor Create(const aHostName, aProgId: string; aUpdateRate: Integer);
    destructor Destroy; override;
  end;

  TDa1Sink = class(TInterfacedObject, IAdviseSink)
    Group: TVSC;
    procedure OnDataChange(const formatetc: TFormatEtc; const stgmed: TStgMedium);
      stdcall;
    procedure OnViewChange(dwAspect: Longint; lindex: Longint);
      stdcall;
    procedure OnRename(const mk: IMoniker); stdcall;
    procedure OnSave; stdcall;
    procedure OnClose; stdcall;
    constructor Create(aGroup: TVSC);
  end;

  TDa2Sink = class(TInterfacedObject, IOPCDataCallback)
    Group: TVSC;
    function OnDataChange(dwTransid: DWORD; hGroup: OPCHANDLE;
      hrMasterquality, hrMastererror: HResult; dwCount: DWORD;
      phClientItems: POPCHANDLEARRAY; pvValues: POleVariantArray;
      pwQualities: PWORDARRAY; pftTimeStamps: PFileTimeArray;
      pErrors: PResultList): HResult; stdcall;
    function OnReadComplete(dwTransid: DWORD; hGroup: OPCHANDLE;
      hrMasterquality, hrMastererror: HResult; dwCount: DWORD;
      phClientItems: POPCHANDLEARRAY; pvValues: POleVariantArray;
      pwQualities: PWORDARRAY; pftTimeStamps: PFileTimeArray;
      pErrors: PResultList): HResult; stdcall;
    function OnWriteComplete(dwTransid: DWORD; hGroup: OPCHANDLE;
      hrMastererr: HResult; dwCount: DWORD; pClienthandles: POPCHANDLEARRAY;
      pErrors: PResultList): HResult; stdcall;
    function OnCancelComplete(dwTransid: DWORD;
      hGroup: OPCHANDLE): HResult; stdcall;
    constructor Create(aGroup: TVSC);
  end;

  EUnknownId = class(Exception);

  PNameArray = ^TNameArray;
  TNameArray = array[Word] of PWideChar;
  PItemRecArray = ^TItemRecArray;
  TItemRecArray = array[Word] of TItemRec;

{ TServer }

var
  Servers: TStringList;
  Da1DataTimeFormat: Integer = 0;

{$IFDEF Debug}
procedure ShowTicks(const s: string);
begin
  OutputDebugString(PChar(
    Format('%d : %s', [GetTickCount, s])))
end;
{$ENDIF}

function TServer.AddGroup(UpdateRate, ClientHandle: Integer; var ServerHandle: OPCHANDLE): IUnknown;
var
  NullChar: WideChar;
  ActualRate: Cardinal;
begin
  NullChar:= #0;
  OleCheck(Server.AddGroup(@NullChar, UpdateRate > 0, UpdateRate, ClientHandle,
    nil, nil, GetUserDefaultLCID, ServerHandle, ActualRate, IUnknown, Result));
  Inc(GroupCount);
end;

class function TServer.CreateServer(const aHostName,
  aProgId: string): TServer;
var
  i: Integer;
  Ndx: string;
  Clsid: TGUID;
{ CPC: IConnectionPointContainer;
  CP: IConnectionPoint; }
begin
  if not Assigned(Servers) then
    Servers:= TStringList.Create;
  Ndx:= aHostName + '.' + aProgId;
  if Servers.Find(Ndx, i) then
  begin
    Result:= TServer(Servers.Objects[i])
  end else
  begin
    if aProgID = '' then
      raise EOleSysError.CreateRes(@SNoProgId);
    Result:= TServer.Create;
    try
      with Result do
      begin
        GetServerInfo(aHostName, aProgID, Clsid, ServerDesc, VendorName);
        if IsLocalHost(aHostName) then
          Server:= CreateComObject(Clsid) as IOPCServer
        else
          Server:= CreateRemoteComObject(aHostName, Clsid) as IOPCServer;
        ShutdownCookie:= 0;
        { Have not quite decided how to implement this yet
        if (Server.QueryInterface(IConnectionPointContainer, CPC) = S_OK) and
           (CPC.FindConnectionPoint(IOpcShutdown, CP) = S_OK) then
          CP.Advise(TShutdownSink.Create(Result), ShutdownCookie)
        }
      end;
      Servers.AddObject(Ndx, Result)
    except
      Result.Free;
      raise
    end
  end
end;

destructor TServer.Destroy;
var
  i: Integer;
begin
  if Assigned(Servers) then
  begin
    i:= Servers.IndexOfObject(Self);
    if i <> -1 then
    begin
      Servers.Delete(i);
      if Servers.Count = 0 then
        FreeAndNil(Servers)
    end
  end;
  inherited Destroy
end;

function OpcClient(
  const HostName, ProgId: string;
  UpdateRate: Integer): IOpcClient;
begin
  Result:= TVSC.Create(HostName, ProgId, UpdateRate) as IOpcClient
end;

procedure TServer.RemoveGroup(hGroup: OPCHANDLE);
begin
  Server.RemoveGroup(hGroup, false);
  Dec(GroupCount);
  if GroupCount = 0 then
    Destroy
end;

(*
{ TShutdownSink }

constructor TShutdownSink.Create(aServer: TServer);
begin
  inherited Create;
  Server:= aServer
end;

function TShutdownSink.ShutdownRequest(szReason: POleStr): HResult;
begin
  {&&& what to do about this?}
end;
*)

procedure TVSC.DA1DataTimeChange(dwCount: DWORD; phClientHeaders: POPCITEMHEADER1ARRAY;
                Values: DAVariant);
var
  i: Integer;
  ItemRec: TItemRec;
begin
  for i:= 0 to dwCount - 1 do
  with phClientHeaders^[i] do
  begin
    ItemRec:= TItemRec(hClient);
    ItemRec.Value:= Values[i];
    ItemRec.Quality:= wQuality;
    ItemRec.Timestamp:= FileTimeToDateTime(ftTimestampItem)
  end
end;

constructor TVSC.Create(const aHostName, aProgId: string; aUpdateRate: Integer);
{if aUpdateRate = 0 then no callbacks}
var
  CPC: IConnectionPointContainer;
  CP: IConnectionPoint;
  DataObj: IDataObject;
  FormatEtc: TFormatEtc;
  DA1Sink: IUnknown;
begin
  inherited Create;
  Server:= TServer.CreateServer(aHostName, aProgId);
  UpdateRate:= aUpdateRate;
  {add Group}
  Group:= Server.AddGroup(UpdateRate, Integer(Self), hGroup);
  if UpdateRate > 0 then  {attempt to connect}
  begin
    if Group.QueryInterface(IConnectionPointContainer, CPC) = S_OK then
    begin
      OleCheck(CPC.FindConnectionPoint(IOpcDataCallback, CP));
      CP.Advise(TDa2Sink.Create(Self), Da2Cookie)
    end else
    if Group.QueryInterface(IDataObject, DataObj) = S_OK then
    begin
      if Da1DataTimeFormat = 0 then
        Da1DataTimeFormat:= RegisterClipboardFormat('OPCSTMFORMATDATATIME');
      with FormatEtc do
      begin
        cfFormat:= Da1DataTimeFormat;
        dwAspect:= DVASPECT_CONTENT;
        ptd:= nil;
        tymed:= TYMED_HGLOBAL;
        lindex:= -1
      end;
      DA1Sink:= TDA1Sink.Create(Self);
      OleCheck(DataObj.DAdvise(FormatEtc, ADVF_PRIMEFIRST, DA1Sink as IAdviseSink, Da1Cookie))
    end
  end;
  Items:= TStringList.Create;
  Items.Sorted:= true;
  Items.Duplicates:= dupError
end;

destructor TVSC.Destroy;
var
  i: Integer;
  CPC: IConnectionPointContainer;
  CP: IConnectionPoint;
  DataObj: IDataObject;
begin
  if Assigned(Group) then
  begin
    if DA2Cookie <> 0 then
    begin
      CPC:= Group as IConnectionPointContainer;
      OleCheck(CPC.FindConnectionPoint(IOpcDataCallback, CP));
      CP.Unadvise(DA2Cookie);
      DA2Cookie:= 0
    end else
    if Da1Cookie <> 0 then
    begin
      DataObj:= Group as IDataObject;
      DataObj.DUnadvise(Da1Cookie);
      Da1Cookie:= 0
    end
  end;
  Group:= nil;
  if Assigned(Server) then
    Server.RemoveGroup(hGroup);
  if Assigned(Items) then
  begin
    for i:= 0 to Items.Count - 1 do
      Items.Objects[i].Free;
    Items.Destroy
  end;
  inherited Destroy
end;

function TVSC.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
var
  NameArray: PNameArray absolute Names;
  ItemRecArray: PItemRecArray absolute DispIDs;
  i: Integer;
begin
  for i:= 0 to NameCount - 1 do
    ItemRecArray[i]:= GetDispId(NameArray[i]);
  Result:= S_OK
end;

function TVSC.GetTypeInfo(Index, LocaleID: Integer;
  out TypeInfo): HResult;
begin
  Result:= E_FAIL
end;

function TVSC.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Count:= 0;
  Result:= S_OK
end;

function TVSC.Invoke(DispID: Integer; const IID: TGUID;
  LocaleID: Integer; Flags: Word; var Params; VarResult, ExcepInfo,
  ArgErr: Pointer): HResult;
var
  DispParams: TDispParams absolute Params;
  ItemRec: TItemRec;
begin
  try
    ItemRec:= TItemRec(DispID);
    if (Flags and DISPATCH_PROPERTYGET) <> 0 then
    begin
      OleVariant(VarResult^):= ItemRec.Get;
      Result:= S_OK
    end else
    if (Flags and DISPATCH_PROPERTYPUT) <> 0 then
    begin
      if DispParams.cArgs = 1 then
      begin
        ItemRec.Put(POleVariant(DispParams.rgvarg)^);
        Result:= S_OK
      end else
      begin
        Result:= DISP_E_BADPARAMCOUNT
      end
    end else
    begin
      Result:= E_FAIL
    end
  except
    on E: EOleSysError do
      Result:= E.ErrorCode
  end;
end;

function TVSC.GetDispId(Name: PWideChar): TItemRec;
begin
  Result:= GetItemRec(Name) {do this to ensure only one conversion wide->ansi}
end;

function TVSC.GetItemRec(const Name: string): TItemRec;
{note: DispId is a pointer to an item record}
var
  i: Integer;
  NewItem: TItemRec;
begin
  {add a new item to the list if necessary}
  if Items.Find(Name, i) then
  begin
    Result:= TItemRec(Items.Objects[i])
  end else
  begin
    try
      NewItem:= TItemRec.Create(Self, Name);
      Items.AddObject(Name, NewItem);
      Result:= NewItem
    except
      on EUnknownId do
        Result:= Pointer(DISP_E_UNKNOWNNAME)
    end
  end
end;

{ TItemRec }

constructor TItemRec.Create(aOwner: TVSC; const ItemName: string);
var
  ItemDef: OPCITEMDEF;
  ItemRes: POPCITEMRESULTARRAY;
  WideName: WideString;
  Results: PResultList;
  Res: HRESULT;
  AR: TAccessRights;
  IAR: TItemAccessRights absolute AR;
begin
  inherited Create;
  Owner:= aOwner;
  FillChar(ItemDef, SizeOf(ItemDef), 0);
  WideName:= ItemName;
  with ItemDef do
  begin
    szItemID:= PWideChar(WideName);
    bActive:= true;
    hClient:= OPCHANDLE(Self)
  end;
  ItemRes:= nil;
  Results:= nil;
  try
    with Owner.Group as IOpcItemMgt do
      Res:= AddItems(1, @ItemDef, ItemRes, Results);
    if Res = OPC_E_INVALIDITEMID then
      raise EUnknownId.Create('');
    OleCheck(Res);
    hServer:= ItemRes^[0].hServer;
    CanonicalDataType:= ItemRes^[0].vtCanonicalDataType;
    AR:= NativeAccessRights(ItemRes^[0].dwAccessRights);
    {types AccessRights and ItemAccessRights are identical}
    AccessRights:= IAR
  finally
    FreeOpcItemResultArray(1, ItemRes);
    FreeResultList(Results)
  end
end;

function TItemRec.Get: OleVariant;
var
  Errors: PResultList;
  Values: POPCITEMSTATEARRAY;
begin
  if Owner.UpdateRate = 0 then
  begin  {SyncIO Read}
    Values:= nil;
    Errors:= nil;
    try
      with Owner.Group as IOPCSyncIO do
        OleCheck(Read(OPC_DS_DEVICE, 1, @(hServer), Values, Errors));
      Timestamp:= FiletimeToDateTime(Values^[0].ftTimeStamp);
      Quality:= Values^[0].wQuality;
      Value:= Values^[0].vDataValue;
      Result:= Value
    finally
      FreeOpcItemStateArray(1, Values);
      FreeResultList(Errors)
    end
  end else
  begin
    Result:= Value
  end
end;

procedure TItemRec.Put(const NewValue: OleVariant);
var
  Errors: PResultList;
begin
  Errors:= nil;
  try
    with Owner.Group as IOPCSyncIO do
      OleCheck(Write(1, @(hServer), @NewValue, Errors));
    if Owner.UpdateRate = 0 then
    begin
      Value:= NewValue;
      Timestamp:= Now
    end
  finally
    FreeResultList(Errors)
  end;
end;

{ TDa1Sink }

constructor TDa1Sink.Create(aGroup: TVSC);
begin
  inherited Create;
  Group:= aGroup
end;

procedure TDa1Sink.OnClose;
begin

end;

procedure TDa1Sink.OnDataChange(const formatetc: TFormatEtc;
  const stgmed: TStgMedium);
var
  GroupHeader: POPCGROUPHEADER;
  DataTimeItemHeaders: POPCITEMHEADER1ARRAY;
  Data: DAVariant;
begin
  if formatetc.cfFormat = DA1DataTimeFormat then
  begin
    FormatOPCStmDataTime( GlobalLock(stgmed.hGlobal),
                         GroupHeader, DataTimeItemHeaders, Data);
    try
      Group.DA1DataTimeChange(GroupHeader^.dwItemCount, DataTimeItemHeaders, Data)
    finally
      GlobalUnlock(stgmed.hGlobal)
    end
  end
end;

procedure TDa1Sink.OnRename(const mk: IMoniker);
begin
end;

procedure TDa1Sink.OnSave;
begin
end;

procedure TDa1Sink.OnViewChange(dwAspect, lindex: Integer);
begin
end;

{ TDa2Sink }

constructor TDa2Sink.Create(aGroup: TVSC);
begin
  inherited Create;
  Group:= aGroup
end;

function TDa2Sink.OnCancelComplete(dwTransid: DWORD;
  hGroup: OPCHANDLE): HResult;
begin
  Result:= S_OK
end;

function TDa2Sink.OnDataChange(dwTransid: DWORD; hGroup: OPCHANDLE;
  hrMasterquality, hrMastererror: HResult; dwCount: DWORD;
  phClientItems: POPCHANDLEARRAY; pvValues: POleVariantArray;
  pwQualities: PWORDARRAY; pftTimeStamps: PFileTimeArray;
  pErrors: PResultList): HResult;
begin
  if hGroup = DWORD(Group) then
    Group.DataChange(dwCount, phClientItems, pvValues,
      pwQualities, pftTimestamps, pErrors);
  Result:= S_OK
end;

function TDa2Sink.OnReadComplete(dwTransid: DWORD; hGroup: OPCHANDLE;
  hrMasterquality, hrMastererror: HResult; dwCount: DWORD;
  phClientItems: POPCHANDLEARRAY; pvValues: POleVariantArray;
  pwQualities: PWORDARRAY; pftTimeStamps: PFileTimeArray;
  pErrors: PResultList): HResult;
begin
  Result:= S_OK
end;

function TDa2Sink.OnWriteComplete(dwTransid: DWORD; hGroup: OPCHANDLE;
  hrMastererr: HResult; dwCount: DWORD; pClienthandles: POPCHANDLEARRAY;
  pErrors: PResultList): HResult;
begin
  Result:= S_OK
end;

procedure TVSC.DataChange(dwCount: DWORD; phClientItems: POPCHANDLEARRAY;
  pvValues: POleVariantArray; pwQualities: PWORDARRAY;
  pftTimeStamps: PFileTimeArray; pErrors: PResultList);
var
  i: Integer;
  Item: TItemRec;
begin
  for i:= 0 to dwCount - 1 do
  begin
    Item:= TItemRec(phClientItems^[i]);
    if Assigned(pftTimestamps) then
      Item.Timestamp:= FileTimeToDateTime(pftTimestamps^[i])
    else
      Item.Timestamp:= Now;
    if Assigned(pwQualities) then
      Item.Quality:= pwQualities^[i]
    else
      Item.Quality:= OPC_QUALITY_GOOD;
    Item.Value:= pvValues^[i]
  end
end;

function TVSC.GetAccessRights(const Name: string): TItemAccessRights;
begin
  Result:= GetItemRec(Name).AccessRights
end;

function TVSC.GetItemType(const Name: string): TVarType;
begin
  Result:= GetItemRec(Name).CanonicalDataType
end;

function TVSC.GetItemValue(const Name: string): OleVariant;
begin
  Result:= GetItemRec(Name).Get
end;

function TVSC.GetPercentDeadband: Single;
begin
  Result:= PercentDeadband
end;

function TVSC.GetUpdateRate: Integer;
begin
  Result:= UpdateRate
end;

procedure TVSC.SetItemValue(const Name: string; const Value: OleVariant);
begin
  GetItemRec(Name).Put(Value)
end;

procedure TVSC.SetPercentDeadband(const Value: Single);
var
  RevisedUpdateRate: DWORD;
begin
  with Group as IOPCGroupStateMgt do
    OleCheck(SetState(nil, RevisedUpdateRate, nil, nil, @Value, nil, nil));
  PercentDeadband:= Value
end;

function TVSC.GetItemQuality(const Name: string): DWORD;
begin
  Result:= GetItemRec(Name).Quality
end;

function TVSC.GetItemTimestamp(const Name: string): TDateTime;
begin
  Result:= GetItemRec(Name).Timestamp
end;

{ EOpcQuality }

constructor EOpcQuality.Create(aQuality: DWORD);
begin
  inherited CreateResFmt(@SBadQuality, [aQuality]);
  Quality:= aQuality
end;

end.
