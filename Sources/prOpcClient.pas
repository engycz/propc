{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
{History
  13-03-01. Dropped PathChar property. I can see real problems with this,
            and I think it is very unlikely to be ever needed.
  14-03-10. Numerous mods to handle new exception scheme.


Release 1.11
1.11.4      OutputDebugString accidentally left in DA2DataChange

Release 1.14
1.14.1      TOpcGroup.FItems was sorted. I did this to make searches more
            efficient and detection of duplicates more efficient but the
            effect was to make item indexes pretty well impossible to track.
            This change may upset some existing applications. To restore
            'old' behaviour, set TOpcSimpleClient.SortedItemLists:= true

1.14.2      Fixed memory leak in SyncReadItem

1.14.3      Added support for getting EUInfo. This is not dynamic: EUInfo is
            assumed to be static for the lifetime of the group. New methods on
            TOpcGroup are:

            function TOpcGroupItemEnumeratedNames(
              i: Integer; Names: TStrings): Boolean;
            function TOpcGroup.ItemAnalogRange(
              i: Integer; var Low, High: Double): Boolean;

            Both return false if the appropriate EU information is not
            available.

Release 1.15f

1.15f.1
add property TOpcGroup.Timebias. This is updated from the server when the
group is connected.

1.15f.2
add property TOpcGroup.GroupInterface to allow access to underlying
OpcInterfaces.

1.15h.1
Modify 'DisconnectServer' to avoid exception if connection point cannot
be found (e.g Server has unexpectedly died)


1.16a.1
Add support for AsyncIO operations

1.16c.1 (BETA)
Eliminate use of FItemList.Objects. Test build for Markus Muller.


1.16e.1
Use OPCENUM.EXE when opening connection at runtime.
}

unit prOpcClient;
{$I prOpcCompilerDirectives.inc}
interface

uses
  Windows, Messages, SysUtils, Classes, prOpcComn, prOpcDa,
  prOpcError, prOpcUtils, ActiveX, prOpcTypes, prOpcEnum;
  {try to keep any VCL stuff out of here - i.e Controls, Forms}

const
  WM_OPCBROWSEUPDATE = WM_USER + 348;

type
  TOpcSimpleClient = class;
  TOpcGroup = class;

  TDataChangeEvent = procedure(Sender: TOpcGroup; ItemIndex: Integer; const NewValue: Variant;
                                NewQuality: Word; NewTimestamp: TDateTime) of object;

  TWriteCompleteEvent = procedure(Sender: TOpcGroup; ItemIndex: Integer; Result: HRESULT) of object;

  TBrowseNodeEvent = function (Parent: Pointer; const BrowseId, ItemId: string): Pointer of object;

  TItemProperty = record
    ID: DWORD;
    Desc: string;
    Datatype: Integer;
    Value: OleVariant;
    Error: HRESULT;
  end;

  TServerStatus = record
    StartTime:      TDateTime;
    CurrentTime:    TDateTime;
    LastUpdateTime: TDateTime;
    ServerState:    OPCSERVERSTATE;
    GroupCount:     DWORD;
    BandWidth:      DWORD;
    MajorVersion:   Word;
    MinorVersion:   Word;
    BuildNumber:    Word;
    Reserved:       Word;
    VendorInfo:     string;
  end;

  TItemProperties = array of TItemProperty;

  TOpcGroup =
  class(TCollectionItem)
  private
    FName: string;
    FServerAssignedName: string;
    FItems: TStringList;
    FItemInfo: Pointer;  {cf 1.16c.2}
    FItemInfoCount : Word;
    FUpdateRate: Cardinal;
    FActive: Boolean;
    FPercentDeadband: Single;
    FTimebias: Longint; {cf 1.15f.1}
    FServerHandle: OPCHANDLE;
    FOpcGroup: IUnknown;
    FDataAccessType: TOpcDataAccessType;
    FDa2Cookie: Longint;
    FDa1Cookie: Longint;

    FAccessPath: string;
    FUpdateList: TList;
    FUpdating: (umNone, umSync, umASync);
    FOnDataChange: TDataChangeEvent;
    FOnWriteComplete: TWriteCompleteEvent;
    procedure SetItems(Value: TStrings);
    function GetItems: TStrings;
    procedure SetUpdateRate(Value: Cardinal);
    procedure SetActive(Value: Boolean);
    procedure SetPercentDeadband(Value: Single);
    procedure SetAccessPath(const Value: String); {not while connected}
    procedure GetItemDef(i: Integer; var ItemDef: TOpcItemDef);
    procedure FreeItemInfo;
    procedure CheckNotConnected;
    procedure CheckConnected;
    procedure CheckAdvise;
    procedure CheckActive;
    function GetItemLastResult(i: Integer): HRESULT;
    procedure SetItemLastResult(i: Integer; Value: HRESULT);
    function GetItemValue(i: Integer): OleVariant; {cache}
    procedure SetItemValue(i: Integer; const Value: OleVariant); {cache}
    function GetItemAccessRights(i: Integer): TAccessRights;
    function GetItemCanonicalDataType(i: Integer): TVarType;
    function GetItemAccessPath(i: Integer): String;
    function GetItemQuality(i: Integer): Word;
    function GetItemTimestamp(i: Integer): TDateTime;
    function GetItemActive(i: Integer): Boolean;
    function GetItemID(i: Integer): String;
    function GetServerHandle(i: Integer): OPCHANDLE;
    function GetDataAccessType: TOpcDataAccessType;
    procedure SetItemActive(i: Integer; const Value: Boolean);
    procedure ClearUpdateList;
    procedure DataChange(dwCount: DWORD; phClientItems: POPCHANDLEARRAY;
                pvValues: POleVariantArray; pwQualities: PWORDARRAY;
                pftTimeStamps: PFileTimeArray; pErrors: PResultList);
    procedure DA1DataTimeChange(dwCount: DWORD; phClientHeaders: POPCITEMHEADER1ARRAY;
                Values: DAVariant);
    procedure WriteComplete(dwCount: DWORD; phClientItems: POPCHANDLEARRAY; pErrors: PResultList);
    procedure DisconnectServer(const Server: IOpcServer);
    procedure ConnectServer(const Server: IOpcServer);
    procedure SetName(const Value: String);
    function GetSortedItemList: Boolean;
    procedure SetSortedItemList(const Value: Boolean);
    procedure DoEndUpdate;
    procedure DoWriteItem(i: Integer; const Value: OleVariant; ASync: Boolean);
  protected
    function GetDisplayName: String; override;
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
    procedure Assign( Source: TPersistent); override;
    procedure BeginUpdate;
    procedure EndUpdate; {flush all reads/writes since BeginUpdate}
    procedure AsyncBeginUpdate;
    procedure AsyncEndUpdate; {flush all reads/writes since BeginUpdate}
    function SyncReadItem(i: Integer): OleVariant;
    function IsConnected: Boolean;
    procedure SyncWriteItem(i: Integer; const Value: OleVariant);
    procedure ASyncWriteItem(i: Integer; const Value: OleVariant);
    procedure ItemCheckAccessRights(i: Integer; RequiredRights: TAccessRights);
    procedure Refresh; {from server}
    procedure OpcCheck(Res: HRESULT);
    function Client: TOpcSimpleClient;
    procedure Connect;
    procedure Disconnect;
    function ItemEnumeratedNames(i: Integer; Names: TStrings): Boolean;
    function ItemAnalogRange(i: Integer; var Low, High: Double): Boolean;
    property GroupInterface: IUnknown read FOpcGroup; {cf 1.15f.1}
    property Timebias: Longint read FTimebias; {cf 1.15f.1}
    property SortedItemList: Boolean read GetSortedItemList write SetSortedItemList;
    property ServerAssignedName: string read FServerAssignedName;
    property DataAccessType: TOpcDataAccessType read GetDataAccessType;
    property ItemValue[i: Integer]: OleVariant read GetItemValue write SyncWriteItem;
    property ItemAccessPath[i: Integer]: String read GetItemAccessPath;
    property ItemID[i: Integer]: String read GetItemID;
    property ItemAccessRights[i: Integer]: TAccessRights read GetItemAccessRights;
    property ItemCanonicalDataType[i: Integer]: TVarType read GetItemCanonicalDataType;
    property ItemActive[i: Integer]: Boolean read GetItemActive write SetItemActive;
    property ItemLastResult[i: Integer]: HRESULT read GetItemLastResult;
    property ItemQuality[i: Integer]: Word read GetItemQuality;
    property ItemTimestamp[i: Integer]: TDateTime read GetItemTimestamp;
  published
    property Name: String read FName write SetName;
    property UpdateRate: Cardinal read FUpdateRate write SetUpdateRate;
    property Active: Boolean read FActive write SetActive;
    property PercentDeadband: Single read FPercentDeadband write SetPercentDeadband;
    property Items: TStrings read GetItems write SetItems;
    property AccessPath: String read FAccessPath write SetAccessPath; {default if items do not have accesspath}
    property OnDataChange: TDataChangeEvent read FOnDataChange write FOnDataChange;
    property OnWriteComplete: TWriteCompleteEvent read FOnWriteComplete write FOnWriteComplete;
  end;

  TOpcGroupCollection =
  class(TCollection)
  private
    FOpcClient: TOpcSimpleClient;
    function GetGroup(i: Integer): TOpcGroup;
    procedure SetGroup(i: Integer; Value: TOpcGroup);
    procedure ServerAddGroups(const Server: IOpcServer);
    procedure ServerRemoveGroups(const Server: IOpcServer);
  protected
    function GetOwner: TPersistent; override;
  public
    function Add: TOpcGroup;
    constructor Create(aOpcClient: TOpcSimpleClient);
    property Groups[i: Integer]: TOpcGroup read GetGroup write SetGroup; default;
    property Client: TOpcSimpleClient read FOpcClient;
  end;

  TOpcSimpleClient = class(TComponent)
  private
    FGroups: TOpcGroupCollection;
    FOpcServer: IOpcServer;
    FHostName: String;
    FProgID: String;
    FServerDesc: String;
    FVendorName: String;
    FActive: Boolean;
    FOnConnect: TNotifyEvent;
    FOnDisconnect: TNotifyEvent;
    FBrowserWindow: TObject;
    FShutdownReason: string;
    FConnectIOPCShutdown: Boolean;
    FShutdownCookie: Integer;
    FOnServerShutdown: TNotifyEvent;
    FSortedItemLists: Boolean;
    function GetActive: Boolean;
    procedure SetActive(Value: Boolean);
    procedure SetHostName(const Value: String);
    procedure SetProgID(const Value: String);
    procedure CheckActive;
    procedure CheckNotActive;
    procedure UpdateBrowser;
    procedure ServerShutdown(pzReason: PWideChar);
    procedure GetBrowser(var Browse: IOPCBrowseServerAddressSpace);
    procedure WMOpcClientConnect(var Msg: TMessage); message WM_OPCBROWSEUPDATE;
              {browser calls this to connect. Avoids having public
               procedure for doing the same. This component uses the
               same message ID to cause refresh on browser}
    procedure SetGroups( Value: TOpcGroupCollection);
    procedure SetSortedItemLists(const Value: Boolean);
  protected
{    procedure Loaded; override;}
  public
    constructor Create(aOwner: TComponent); override;
    destructor Destroy; override;
    procedure GetAllItems( ItemIds: TStrings);
    function ShutdownConnected: Boolean;
    procedure GetItemProperties(const ItemId: string; var Properties: TItemProperties);
    procedure UpdateProperties(const ItemId: string; var Properties: TItemProperties);
    function SupportsBrowsing: Boolean;
    procedure OpcCheck(Res: HRESULT);
    function GetErrorString(Res: HRESULT): string;
    procedure Connect(NoGroups: Boolean = false);
    procedure ConnectGroups;
    procedure Disconnect;
    procedure DisconnectGroups;
    procedure BrowseItems(BrowseNodeEvent: TBrowseNodeEvent); virtual;
    function GetServerStatus: TServerStatus;
    property ServerDesc: String read FServerDesc;
    property VendorName: String read FVendorName;
    property OpcServer: IOpcServer read FOpcServer;
    property ShutdownReason: string read FShutdownReason;
  published
    property Groups: TOpcGroupCollection read FGroups write SetGroups;
    property HostName: String read FHostName write SetHostName;
    property ProgID: String read FProgID write SetProgID;
    property Active: Boolean read GetActive write SetActive stored false;
    property SortedItemLists: Boolean read FSortedItemLists write SetSortedItemLists;
    property ConnectIOPCShutdown: Boolean read FConnectIOPCShutdown write FConnectIOPCShutdown;
    property OnConnect: TNotifyEvent read FOnConnect write FOnConnect;
    property OnDisconnect: TNotifyEvent read FOnDisconnect write FOnDisconnect;
    property OnServerShutdown: TNotifyEvent read FOnServerShutdown write FOnServerShutdown;
  end;

  EOpcClient = class(Exception)
    constructor CreateHResult(aClient: TOpcSimpleClient; Res: HRESULT);
  end;

implementation
uses
  {$IFDEF D6UP}
  Variants,
  RTLConsts,
  {$ELSE}
  Consts,
  {$ENDIF}
  ComObj;

const
  MinPercentDeadband = 0.01;

resourcestring
  SNotActiveServer = 'Operation not allowed on active server';
  SNotActiveGroup = 'Operation not allowed on active group';
  SNoProgID = 'Server ProgID must be specified';
  SCannotAddGroupTwice = 'Cannot add group to server twice';
  SMustBeConnected = 'Group must be connected';
  SBadAccessRights = 'Insufficient access rights';
  SRequiresAdvise = 'Operation requires advise group';
  SRequiresActiveGroup = 'Operation requires active group';
  SRequiresActiveServer = 'Operation requires active server';
  SNoAdvise = 'No active advise (Unexpected)';
  SIOPCBrowseNotSupported = 'Server does not support IOPCBrowseServerAddressSpace';
  SDuplicateGroupName = 'Duplicate group name not allowed';
  SCannotFindClient = 'Group cannot find client';
  SNoOpcServer = 'Opc server is not assigned';
  SBadlyFormedEUInfo = 'Server provided EUInfo incorrect type';

var
  Da1DataTimeFormat: Longint = 0;
{we only support DataTime streams. This might mean we fail to
connect with a server that supports only Data streams (this
would be non-compliant. To support DataTime streams would not
be such a big deal. (add a property?)}


type
  TDa1Sink = class(TInterfacedObject, IAdviseSink)
    FGroup: TOpcGroup;
    procedure OnDataChange(const formatetc: TFormatEtc; const stgmed: TStgMedium);
      stdcall;
    procedure OnViewChange(dwAspect: Longint; lindex: Longint);
      stdcall;
    procedure OnRename(const mk: IMoniker); stdcall;
    procedure OnSave; stdcall;
    procedure OnClose; stdcall;
    constructor Create(aGroup: TOpcGroup);
  end;

  TDa2Sink = class(TInterfacedObject, IOPCDataCallback)
    FGroup: TOpcGroup;
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
    constructor Create(aGroup: TOpcGroup);
  end;

  TShutdownSink = class(TInterfacedObject, IOPCShutdown)
    FClient: TOpcSimpleClient;
    function ShutdownRequest(szReason: POleStr): HResult; stdcall;
    constructor Create(aClient: TOpcSimpleClient);
  end;

  PItemInfo = ^TItemInfo;
  TItemInfo = record
    LastResult: HRESULT;
    AccessPath: String;
    ItemID: String;
    Active: Boolean;
    hClient: OPCHANDLE;
    hServer: OPCHANDLE;
    AccessRights: TAccessRights;
    EUType: TEuType;
    EUInfo: OleVariant;
    EUInfoChecked: Boolean; {only check EUInfo if the client cares}
    CanonicalDataType: TVarType;
    Value: OleVariant;
    Timestamp: TDateTime;
    Quality: Word;
  end;

  {cf 1.16c.1}
  TItemInfoList = array[Word] of TItemInfo;
  PItemInfoList = ^TItemInfoList;

  PUpdateInfo = ^TUpdateInfo;
  TUpdateInfo = record
    Index: Integer;
    Value: OleVariant;
  end;

{I don't want to make TItemInfo public}
function GetItemInfo(Group: TOpcGroup; i: Integer): PItemInfo;
begin
  if (i < 0) or (i >= Group.FItemInfoCount) then
    raise EStringListError.CreateResFmt(@SListIndexError, [i]);
  Result:= @(PItemInfoList(Group.FItemInfo)^[i])
end;


procedure CheckEUInfo(var Info: TItemInfo);

  procedure BadlyFormed;
  begin
    raise EOpcClient.CreateRes(@SBadlyFormedEUInfo)
  end;

begin
  with Info do
  if not EUInfoChecked then
  begin
    case EUType of
      euEnumerated:
      if (VarType(EUInfo) <> (varOleStr or varArray)) or
        (VarArrayDimCount(EUInfo) <> 1) then
        raise EOpcClient.CreateRes(@SBadlyFormedEUInfo);
      euAnalog:
      if (VarType(EUInfo) <> (varDouble or varArray)) or
         (VarArrayDimCount(EUInfo) <> 1) or
         ((VarArrayHighBound(EUInfo, 1) - VarArrayLowBound(EUInfo, 1)) <> 1) then
        BadlyFormed;
    end;
    Info.EUInfoChecked:= true
  end
end;

{ TOpcGroup }

procedure TOpcGroup.Assign(Source: TPersistent);
var
  Src: TOpcGroup absolute Source;
begin
  if Source is TOpcGroup then
  begin
    Name:= Src.Name;
    Items.Assign(Src.Items)
  end else
  begin
    inherited Assign(Source)
  end
end;

procedure TOpcGroup.ClearUpdateList;
var
  i: Integer;
begin
  for i:= 0 to FUpdateList.Count - 1 do
    Dispose(PUpdateInfo(FUpdateList[i]));
  FUpdateList.Clear
end;

procedure TOpcGroup.BeginUpdate;
begin
  ClearUpdateList;
  FUpdating:= umSync
end;

procedure TOpcGroup.AsyncBeginUpdate;
begin
  ClearUpdateList;
  FUpdating:= umAsync
end;

constructor TOpcGroup.Create(Collection: TCollection);
begin
  inherited Create(Collection);
  FItems:= TStringList.Create;
  FItems.Sorted:= Client.SortedItemLists;            {cf 1.14.1}
  FItems.Duplicates:= dupError;     {no effect on Unsorted list}
  FUpdateList:= TList.Create
end;

destructor TOpcGroup.Destroy;
begin
  Disconnect;
  ClearUpdateList;
  FUpdateList.Free;
  FItems.Free;
  inherited
end;

procedure TOpcGroup.DoEndUpdate;
var
  i: Integer;
  hServer: array of OPCHANDLE;
  Results: PResultList;
  Values: array of OleVariant;
  ItemCount: DWORD;
  Async: Boolean;
  CancelID: DWORD;
begin
  Async:= FUpdating = umAsync;
  FUpdating:= umNone;
  ItemCount:= FUpdateList.Count;
  if ItemCount > 0 then
  begin
    SetLength(hServer, ItemCount);
    SetLength(Values, ItemCount);
    for i:= 0 to ItemCount - 1 do
    with PUpdateInfo(FUpdateList[i])^ do
    begin
      hServer[i]:= GetServerHandle(Index);
      Values[i]:= Value
    end;
    Results:= nil;
    try
      if Async then
      begin
        {cf 1.16a.1. Very simple implementation to start. Ignore cancel or transaction id}
        with FOpcGroup as IOpcAsyncIO2 do
          OpcCheck(Write(ItemCount, Pointer(hServer), Pointer(Values), 1, CancelID, Results))
      end else
      begin
        with FOpcGroup as IOpcSyncIO do
          OpcCheck(Write(ItemCount, Pointer(hServer), Pointer(Values), Results))
      end;
      for i:= 0 to ItemCount - 1 do
      with PUpdateInfo(FUpdateList[i])^ do
      begin
        if Assigned(Results) and not Succeeded(Results^[i]) then
        begin
          SetItemLastResult(Index, Results^[i])
        end else
        begin
          SetItemLastResult(Index, S_OK);
          if not Active then
            SetItemValue(Index, Value)
        end
      end
    finally
      FreeResultList(Results)
    end;
    ClearUpdateList
  end
end;

procedure TOpcGroup.EndUpdate;
begin
  DoEndUpdate
end;

procedure TOpcGroup.AsyncEndUpdate;
begin
  DoEndUpdate
end;

function TOpcGroup.GetDisplayName: String;
begin
  if FName = '' then
    Result:= inherited GetDisplayName
  else
    Result:= FName
end;

procedure TOpcGroup.ConnectServer(const Server: IOpcServer);
var
  pPercentDeadband: PSingle;
  WideName: WideString;
  Res: HResult;
  CPC: IConnectionPointContainer;
  CP: IConnectionPoint;
  DataObj: IDataObject;
  FormatEtc: TFormatEtc;
  ItemCount: Cardinal;
  AddItemDefs: TOpcItemDefArray;
  AddResults: POPCITEMRESULTARRAY;
  EnumAttributes: IUnknown;
  Attributes: POPCITEMATTRIBUTESARRAY;
  Fetched: ULONG;
  AddErrors: PResultList;
  NewItemInfo: PItemInfoList;

  {get state params}
  pUpdateRate: DWORD;
  pActive: BOOL;
  ppName: POleStr;
  pTimeBias: Longint;
  _pPercentDeadband: Single;
  pLCID: TLCID;
  phClientGroup: OPCHANDLE;
  phServerGroup: OPCHANDLE;
  HR: HResult;
  i: Integer;
  Info: PItemInfo;
  DA1Sink: IUnknown;
  ItemMgt: IOPCItemMgt;

begin
  if not Assigned(Server) then
    raise EOpcClient.CreateRes(@SNoOpcServer);
  if Assigned(FOpcGroup) then
    raise EOpcClient.CreateRes(@SCannotAddGroupTwice);
  if FPercentDeadband < MinPercentDeadband then
    pPercentDeadband:= nil
  else
    pPercentDeadband:= @FPercentDeadband;
  WideName:= FName;
  Res:= Server.AddGroup(PWideChar(WideName), FActive, FUpdateRate,
    Integer(Self), nil, pPercentDeadband, GetUserDefaultLCID, FServerHandle,
      FUpdateRate, IUnknown, FOpcGroup);
  try
    if (Res <> OPC_S_UNSUPPORTEDRATE) then
      OpcCheck(Res);
    {get state management to retrieve name, which may have been set by server}
    with FOpcGroup as IOPCGroupStateMgt do
      GetState(pUpdateRate, pActive, ppName, pTimeBias, _pPercentDeadband,
        pLCID, phClientGroup, phServerGroup);
    {we might check the other parameters here &&&}
    FServerAssignedName:= ppName;
    FTimebias:= pTimebias;
    CoTaskMemFree(ppName);
    {AddItems}
    ItemCount:= FItems.Count;
    if ItemCount > 0 then
    begin
      FreeItemInfo;
      SetLength(AddItemDefs, ItemCount);
      for i:= 0 to ItemCount - 1 do
        GetItemDef(i, AddItemDefs[i]);
      AddResults:= nil;
      AddErrors:= nil;
      ItemMgt:= FOpcGroup as IOPCItemMgt;
      HR:= ItemMgt.AddItems(ItemCount, POPCITEMDEFARRAY(AddItemDefs), AddResults, AddErrors);
      try  {if not succeeded out params should be nil anyway, but why take the risk?}
        if not Succeeded(HR) then
          OpcCheck(HR);
        NewItemInfo:= AllocMem(SizeOf(TItemInfo) * ItemCount);   {cf 1.16c.1}
        FillChar(NewItemInfo^, SizeOf(TItemInfo) * ItemCount, 0);
        FItemInfo:= NewItemInfo;
        FItemInfoCount := ItemCount;
        for i:= 0 to ItemCount - 1 do
        begin
          // New(Info);
          //FillChar(Info^, SizeOf(TItemInfo), 0);
          with NewItemInfo^[i] do
          begin
            LastResult:= AddErrors^[i];
            AccessPath:= AddItemDefs[i].szAccessPath;
            ItemID:= AddItemDefs[i].szItemID;
            Active:= AddItemDefs[i].bActive;
            hClient:= i;
            hServer:= AddResults^[i].hServer;
            AccessRights:= NativeAccessRights(AddResults^[i].dwAccessRights);
            CanonicalDataType:= AddResults^[i].vtCanonicalDataType
          end;
          {FItems.Objects[i]:= TObject(Info) cf 1.16c.1}
        end;
      finally
        FreeOpcItemResultArray(ItemCount, AddResults);
        FreeResultList(AddErrors)
      end;
      {now enumerate the group to get the attributes} {cf cf 1.14.3}
      Fetched:= 0;
      ItemMgt.CreateEnumerator(IEnumOPCItemAttributes, EnumAttributes);
      with EnumAttributes as IEnumOPCItemAttributes do
        Next(ItemCount, Attributes, @Fetched);
      try
        for i:= 0 to Fetched - 1 do
        begin
          with Attributes^[i] do
          begin
            {Info:= PItemInfo(FItems.Objects[hClient]);}
            Info:= GetItemInfo(Self, hClient);
            Info^.EUType:= TEuType(dwEUType);
            Info^.EUInfo:= vEUInfo
          end
        end
      finally
        FreeOpcItemAttributesArray(Fetched, Attributes)
      end
    end;
    {is this a DA2 Group?}
    if FOpcGroup.QueryInterface(IConnectionPointContainer, CPC) = S_OK then
    begin
      OleCheck(CPC.FindConnectionPoint(IOpcDataCallback, CP));
      CP.Advise(TDa2Sink.Create(Self), FDa2Cookie);
      FDataAccessType:= opcDa2
    end else
    if FOpcGroup.QueryInterface(IDataObject, DataObj) = S_OK then
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
      {note server ignores ADVF_PRIMEFIRST but behaves as if it were set}
      OleCheck(DataObj.DAdvise(FormatEtc, ADVF_PRIMEFIRST, DA1Sink as IAdviseSink, FDa1Cookie))
    end
  except
    DisconnectServer(Server);
    raise
  end
end;

procedure TOpcGroup.DisconnectServer(const Server: IOpcServer);
var
  CPC: IConnectionPointContainer;
  CP: IConnectionPoint;
  DataObj: IDataObject;
begin
  if not Assigned(Server) then
    raise EOpcClient.CreateRes(@SNoOpcServer);
  if Assigned(FOpcGroup) then
  begin
{    if not Assigned(FOpcGroup) then
      raise EOpcClient.CreateRes(@SCannotRemoveGroup); }
    if FDA2Cookie <> 0 then
    begin
      CPC:= FOpcGroup as IConnectionPointContainer;
      {OleCheck(CPC.FindConnectionPoint(IOpcDataCallback, CP));}
      if Succeeded(CPC.FindConnectionPoint(IOpcDataCallback, CP)) then {1.15h.1}
        CP.Unadvise(FDA2Cookie);
      FDA2Cookie:= 0
    end else
    if FDa1Cookie <> 0 then
    begin
      DataObj:= FOpcGroup as IDataObject;
      DataObj.DUnadvise(FDa1Cookie);
      FDa1Cookie:= 0
    end;
    FreeItemInfo;
    FServerAssignedName:= '';
    FOpcGroup:= nil;
    Server.RemoveGroup(FServerHandle, false)
  end
end;

procedure TOpcGroup.SetActive(Value: Boolean);
var
  bActive: BOOL;
begin
  if FActive <> Value then
  begin
    FActive:= Value;
    if Assigned(FOpcGroup) then
    with FOpcGroup as IOPCGroupStateMgt do
    begin
      bActive:= FActive;
      SetState(nil, FUpdateRate, @bActive, nil, nil, nil, nil)
    end
  end
end;

procedure TOpcGroup.CheckNotConnected;
begin
  if Assigned(FOpcGroup) then {connected}
    raise EOpcClient.CreateRes(@SNotActiveGroup)
end;

procedure TOpcGroup.SetItems(Value: TStrings);
var
  TempStrings: TStringList;
begin
  {cannot alter items while connected}
  CheckNotConnected;
  {check for duplicates cf 1.14.1}
  if not FItems.Sorted then
  begin
    TempStrings:= TStringList.Create;
    try
      TempStrings.Sorted:= true;
      TempStrings.Duplicates:= dupError;
      TempStrings.Assign(Value)
    finally
      TempStrings.Free
    end
  end;
  FItems.Assign(Value)
end;

procedure TOpcGroup.SetName(const Value: String);
var
  WideName: WideString;
  i: Integer;
begin
  if Value <> FName then
  begin
    {Check that the existing name is not already in the group}
    {Note: Group names are case-sensitive}
    if Value <> '' then
    with TOpcGroupCollection(Collection) do
    for i:= 0 to Count - 1 do
      if Groups[i].Name = Value then
        raise EOpcClient.CreateRes(@SDuplicateGroupName);
    {if the new name is '' and the group is connected then just leave the
     existing FServerAssignedName alone. '' means "Don't care" so whatever it is
     is OK}
    if Assigned(FOpcGroup) and (Value <> '') then
    begin
      WideName:= Value;
      with FOpcGroup as IOPCGroupStateMgt do
        OleCheck(SetName(PWideChar(WideName)))
    end;
    FName:= Value;
    SetDisplayName(FName) {notify object inspector that DisplayName has changed}
  end
end;

procedure TOpcGroup.SetPercentDeadband(Value: Single);
begin
  if FPercentDeadband <> Value then
  begin
    FPercentDeadband:= Value;
    if Assigned(FOpcGroup) then
    with FOpcGroup as IOPCGroupStateMgt do
      SetState(nil, FUpdateRate, nil, nil, @FPercentDeadband, nil, nil)
  end
end;

procedure TOpcGroup.SetUpdateRate(Value: Cardinal);
begin
  if FUpdateRate <> Value then
  begin
    if Assigned(FOpcGroup) then
    with FOpcGroup as IOPCGroupStateMgt do
      SetState(@Value, FUpdateRate, nil, nil, nil, nil, nil)
    else
      FUpdateRate:= Value
  end
end;


function TOpcGroup.GetItems: TStrings;
begin
  Result:= FItems
end;

procedure TOpcGroup.FreeItemInfo;
var
  I : Integer;
begin
  for i:= 0 to FItemInfoCount - 1 do
   Finalize(PItemInfoList(FItemInfo)^[i]);

  FreeMem(FItemInfo);
  FItemInfoCount := 0;
  FItemInfo:= nil
  {cf 1.16c.1}
  {for i:= 0 to FItems.Count - 1 do
  if Assigned(FItems.Objects[i]) then
  begin
    Dispose(PItemInfo(FItems.Objects[i]));
    FItems.Objects[i]:= nil
  end }
end;

procedure TOpcGroup.GetItemDef(i: Integer; var ItemDef: TOpcItemDef);
begin
  FillChar(ItemDef, SizeOf(ItemDef), 0);
  with ItemDef do
  begin
    szAccessPath:= FAccessPath;
    szItemId:= FItems[i];
    bActive:= true;
    hClient:= i;
    vtRequestedDataType:= VT_EMPTY;
  end
end;

procedure TOpcGroup.SetAccessPath(const Value: String);
begin
  CheckNotConnected;
  FAccessPath:= Value
end;

function TOpcGroup.GetItemAccessRights(i: Integer): TAccessRights;
begin
  CheckConnected;
  {Result:= PItemInfo(FItems.Objects[i]).AccessRights}
  Result:= GetItemInfo(Self, i)^.AccessRights
end;

function TOpcGroup.GetItemCanonicalDataType(i: Integer): TVarType;
begin
  CheckConnected;
  {Result:= PItemInfo(FItems.Objects[i]).CanonicalDataType}
  Result:= GetItemInfo(Self, i)^.CanonicalDataType
end;

function TOpcGroup.GetItemAccessPath(i: Integer): String;
begin
  CheckConnected;
  {Result:= PItemInfo(FItems.Objects[i]).AccessPath}
  Result:= GetItemInfo(Self, i)^.AccessPath
end;

function TOpcGroup.GetItemActive(i: Integer): Boolean;
begin
  CheckConnected;
  {Result:= PItemInfo(FItems.Objects[i]).Active}
  Result:= GetItemInfo(Self, i)^.Active
end;

function TOpcGroup.GetItemLastResult(i: Integer): HRESULT;
begin
  CheckConnected;
  {Result:= PItemInfo(FItems.Objects[i]).LastResult}
  Result:= GetItemInfo(Self, i)^.LastResult
end;

function TOpcGroup.GetServerHandle(i: Integer): OPCHANDLE;
begin
  {Result:= PItemInfo(FItems.Objects[i]).hServer}
  Result:= GetItemInfo(Self, i)^.hServer
end;

function TOpcGroup.GetItemID(i: Integer): String;
begin
  CheckConnected;
  {Result:= PItemInfo(FItems.Objects[i]).ItemID}
  Result:= GetItemInfo(Self, i)^.ItemID
end;

procedure TOpcGroup.SetItemActive(i: Integer; const Value: Boolean);
var
  hServer: OPCHANDLE;
  Results: PResultList;
begin
  CheckConnected;
  hServer:= GetServerHandle(i);
  Results:= nil;
  with FOpcGroup as IOpcItemMgt do
    OpcCheck(SetActiveState(1, @hServer, Value, Results));
  try
    {results should never be nil if Write succeeds, but
     servers are not to be trusted...}
    if Assigned(Results) and not Succeeded(Results^[0]) then
      OpcCheck(Results^[0]);
  finally
    FreeResultList(Results)
  end
end;

function TOpcGroup.SyncReadItem(i: Integer): OleVariant;
var
  hServer: OPCHANDLE;
  Results: PResultList;
  Values: POPCITEMSTATEARRAY;
begin
  CheckConnected;
  hServer:= GetServerHandle(i);
  Results:= nil;
  Values:= nil; {cf 1.14.2}
  with FOpcGroup as IOpcSyncIO do
    OpcCheck(Read(OPC_DS_DEVICE, 1, @hServer, Values, Results));
  try
    if Assigned(Results) and not Succeeded(Results^[0]) then
    begin
      SetItemLastResult(i, Results^[0]);
      OpcCheck(Results^[0]);
    end else
    begin
      Result:= Values^[0].vDataValue;
      {with PItemInfo(FItems.Objects[i])^ do}  {cf 1.16c.1}
      with GetItemInfo(Self, i)^ do
      begin
        LastResult:= S_OK;
        Value:= Result;
        Timestamp:= FiletimeToDateTime(Values^[0].ftTimestamp);
        Quality:= Values^[0].wQuality
      end
    end
  finally
    FreeResultList(Results);
    FreeOpcItemStateArray(1, Values) {cf 1.14.2}
  end
end;

procedure TOpcGroup.DoWriteItem(i: Integer; const Value: OleVariant; ASync: Boolean);
var
  j: Integer;
  Found: Boolean;
  P: PUpdateInfo;
  hServer: OPCHANDLE;
  Results: PResultList;
  CancelID: DWORD;
begin
  CheckConnected;
  ItemCheckAccessRights(i, [iaWrite]);
  if FUpdating <> umNone then
  begin
    Found:= false;
    for j:= 0 to FUpdateList.Count - 1 do
    begin
      P:= Pointer(FUpdateList[j]);
      if P^.Index = i then
      begin
        P^.Value:= Value;
        Found:= true;
        break
      end
    end;
    if not Found then
    begin
      New(P);
      P^.Index:= i;
      P^.Value:= Value;
      FUpdateList.Add(P)
    end
  end else
  begin
    hServer:= GetServerHandle(i);
    Results:= nil;
    if ASync then
    begin
      with FOpcGroup as IOpcASyncIO2 do
        OpcCheck(Write(1, @hServer, @Value, 2, CancelID, Results))
    end else
    begin
      with FOpcGroup as IOpcSyncIO do
        OpcCheck(Write(1, @hServer, @Value, Results))
    end;
    try
      if Assigned(Results) and not Succeeded(Results^[0]) then
      begin
        SetItemLastResult(i, Results^[0]);
        OpcCheck(Results^[0]);
      end else
      begin
        SetItemLastResult(i, S_OK);
        if not Active then
          SetItemValue(i, Value)
      end
    finally
      FreeResultList(Results)
    end
  end;
end;

procedure TOpcGroup.SyncWriteItem(i: Integer; const Value: OleVariant);
begin
  DoWriteItem(i, Value, false)
end;

procedure TOpcGroup.ASyncWriteItem(i: Integer; const Value: OleVariant);
begin
  DoWriteItem(i, Value, true)
end;

procedure TOpcGroup.CheckConnected;
begin
  if not Assigned(FOpcGroup) then
    raise EOpcClient.CreateRes(@SMustBeConnected)
end;

function TOpcGroup.GetItemValue(i: Integer): OleVariant;
begin
  CheckConnected;
  if not Active then
    Result:= SyncReadItem(i)
  else
    {Result:= PItemInfo(FItems.Objects[i]).Value}
    Result:= GetItemInfo(Self, i)^.Value
end;

procedure TOpcGroup.SetItemLastResult(i: Integer; Value: HRESULT);
begin
  CheckConnected;
  {PItemInfo(FItems.Objects[i]).LastResult:= Value}
  GetItemInfo(Self, i)^.LastResult:= Value
end;

procedure TOpcGroup.SetItemValue(i: Integer; const Value: OleVariant);
begin
  CheckConnected;
  {PItemInfo(FItems.Objects[i]).Value:= Value}
  GetItemInfo(Self, i)^.Value:= Value
end;

procedure TOpcGroup.ItemCheckAccessRights(i: Integer;
  RequiredRights: TAccessRights);
begin
  if not (RequiredRights <= ItemAccessRights[i])  then
    raise EOpcClient.CreateRes(@SBadAccessRights)
end;

procedure TOpcGroup.DataChange(dwCount: DWORD;
  phClientItems: POPCHANDLEARRAY; pvValues: POleVariantArray;
  pwQualities: PWORDARRAY; pftTimeStamps: PFileTimeArray;
  pErrors: PResultList);
var
  i: Integer;
  Index: DWORD;
  NewQuality: Word;
  NewTimestamp: TDateTime;
  Info: PItemInfo;
begin
  for i:= 0 to dwCount - 1 do
  begin
    Index:= phClientItems^[i];
    if Index < DWORD(FItems.Count) then
    begin
      {Info:= PItemInfo(FItems.Objects[Index]);}
      Info:= GetItemInfo(Self, Index);
      if Assigned(Info) then
      begin
        if Assigned(pftTimestamps) then
          NewTimestamp:= FileTimeToDateTime(pftTimestamps^[i])
        else
          NewTimestamp:= Now;
        if Assigned(pwQualities) then
          NewQuality:= pwQualities^[i]
        else
          NewQuality:= OPC_QUALITY_GOOD;
        if Assigned(FOnDataChange) then
          FOnDataChange(Self, Index, pvValues^[i], NewQuality, NewTimestamp);
        Info^.Value:= pvValues^[i];
        Info^.Quality:= NewQuality;
        Info^.Timestamp:= NewTimestamp
      end
    end
  end
end;

procedure TOpcGroup.DA1DataTimeChange(dwCount: DWORD; phClientHeaders: POPCITEMHEADER1ARRAY;
                Values: DAVariant);
var
  i: Integer;
  Index: DWORD;
  Info: PItemInfo;
  NewTimestamp: TDateTime;
begin
  for i:= 0 to dwCount - 1 do
  with phClientHeaders^[i] do
  begin
    Index:= hClient;
    if Index < DWORD(FItems.Count) then
    begin
      {Info:= PItemInfo(FItems.Objects[Index]);}
      Info:= GetItemInfo(Self, Index);
      if Assigned(Info) then
      begin
        NewTimestamp:= FileTimeToDateTime(ftTimestampItem);
        if Assigned(FOnDataChange) then
          FOnDataChange(Self, Index, Values[i], wQuality, NewTimestamp);
        Info^.Value:= Values[i];
        Info^.Quality:= wQuality;
        Info^.Timestamp:= NewTimestamp
      end
    end
  end
end;

procedure TOpcGroup.WriteComplete(dwCount: DWORD; phClientItems: POPCHANDLEARRAY; pErrors: PResultList);
var
  res: HRESULT;
  i: Integer;
begin
  if Assigned(FOnWriteComplete) then
  begin
    for i:= 0 to dwCount - 1 do
    begin
      if Assigned(pErrors) then
        res:= pErrors^[i]
      else
        res:= S_OK;
      FOnWriteComplete(Self, phClientItems^[i], res)
    end
  end
end;

function TOpcGroup.GetItemQuality(i: Integer): Word;
begin
  CheckConnected;
  {Result:= PItemInfo(FItems.Objects[i]).Quality}
  Result:= GetItemInfo(Self, i)^.Quality
end;

function TOpcGroup.GetItemTimestamp(i: Integer): TDateTime;
begin
  CheckConnected;
  {Result:= PItemInfo(FItems.Objects[i]).Timestamp}
  Result:= GetItemInfo(Self, i)^.Timestamp
end;

procedure TOpcGroup.CheckAdvise;
begin
  if not Assigned(FOpcGroup) or not FActive then
    raise EOpcClient.CreateRes(@SRequiresAdvise);
end;

procedure TOpcGroup.Refresh;
{I have not bothered providing support for
cancelling a refresh. Could do later if req &&&}
var
  io2: IOPCAsyncIO2;
  io1: IOPCAsyncIO;
  CancelID: Cardinal; {don't bother saving}
begin
  CheckConnected;
  CheckAdvise;
  if FItems.Count > 0 then {no point otherwise}
  begin
    CheckActive;
    if FDA2Cookie <> 0 then
    begin
      io2:= FOpcGroup as IOPCAsyncIO2;
      io2.Refresh2(OPC_DS_DEVICE, 0, CancelID)
    end else
    if FDA1Cookie <> 0 then
    begin
      io1:= FOpcGroup as IOpcAsyncIO;
      io1.Refresh(FDa1Cookie, OPC_DS_DEVICE, CancelID)
    end else
    begin
      raise EOpcClient.CreateRes(@SNoAdvise)
    end
  end
end;

procedure TOpcGroup.CheckActive;
begin
  if not FActive then
    raise EOpcClient.CreateRes(@SRequiresActiveGroup)
end;

function TOpcGroup.IsConnected: Boolean;
begin
  Result:= Assigned(FOpcGroup)
end;

function TOpcGroup.GetDataAccessType: TOpcDataAccessType;
begin
  CheckConnected;
  Result:= FDataAccessType
end;

procedure TOpcGroup.OpcCheck(Res: HRESULT);
begin
  if not Succeeded(Res) then
    raise EOpcClient.CreateHResult(Client, Res)
end;

function TOpcGroup.Client: TOpcSimpleClient;
begin
  Result:= TOpcSimpleClient(TOpcGroupCollection(Collection).GetOwner);
  if not Assigned(Result) then
    raise EOpcClient.CreateRes(@SCannotFindClient)
end;

procedure TOpcGroup.Connect;
begin
  ConnectServer(Client.OpcServer)
end;

procedure TOpcGroup.Disconnect;
begin
  DisconnectServer(Client.OpcServer)
end;

function TOpcGroup.GetSortedItemList: Boolean;
begin
  Result:= FItems.Sorted
end;

procedure TOpcGroup.SetSortedItemList(const Value: Boolean);
begin
  FItems.Sorted:= Value
end;

function TOpcGroup.ItemEnumeratedNames(i: Integer;
  Names: TStrings): Boolean;
var
  Info: PItemInfo;
  j: Integer;
begin
  CheckConnected;
  {Info:= PItemInfo(FItems.Objects[i]);}
  Info:= GetItemInfo(Self, i);
  Result:= Assigned(Info) and (Info^.EUType = euEnumerated);
  if Result then
  begin
    CheckEUInfo(Info^);
    {it is OK to pass nil in order to find out if the item has enumerated names}
    if Assigned(Names) then
    begin
      Names.BeginUpdate;
      with Info^ do
      try
        for j:= VarArrayLowBound(EUInfo, 1) to VarArrayHighBound(EUInfo, 1) do
          Names.Add(EUInfo[j])
      finally
        Names.EndUpdate
      end
    end
  end
end;

function TOpcGroup.ItemAnalogRange(i: Integer; var Low, High: Double): Boolean;
var
  Info: PItemInfo;
begin
  CheckConnected;
  {Info:= PItemInfo(FItems.Objects[i]);}
  Info:= GetItemInfo(Self, i);
  Result:= Assigned(Info) and (Info^.EUType = euAnalog);
  if Result then
  begin
    CheckEUInfo(Info^);
    with Info^ do
    begin
      Low:= EUInfo[VarArrayLowBound(EUInfo, 1)];
      High:= EUInfo[VarArrayHighBound(EUInfo, 1)];
    end
  end
end;

{ TOpcGroupCollection }

function TOpcGroupCollection.Add: TOpcGroup;
begin
  Result:= TOpcGroup(inherited Add)
end;

constructor TOpcGroupCollection.Create(aOpcClient: TOpcSimpleClient);
begin
  inherited Create(TOpcGroup);
  FOpcClient:= aOpcClient
end;

function TOpcGroupCollection.GetGroup(i: Integer): TOpcGroup;
begin
  Result:= TOpcGroup(inherited GetItem(i))
end;

function TOpcGroupCollection.GetOwner: TPersistent;
begin
  Result:= FOpcClient
end;

procedure TOpcGroupCollection.ServerAddGroups(const Server: IOpcServer);
var
  i: Integer;
begin
  for i:= 0 to Count - 1 do
    Groups[i].ConnectServer(Server)
end;

procedure TOpcGroupCollection.ServerRemoveGroups(const Server: IOpcServer);
var
  i: Integer;
begin
  for i:= 0 to Count - 1 do
    Groups[i].DisconnectServer(Server)
end;

procedure TOpcGroupCollection.SetGroup(i: Integer; Value: TOpcGroup);
begin
  inherited SetItem(i, Value)
end;

{ TOpcSimpleClient }

procedure TOpcSimpleClient.CheckActive;
begin
  if not Active then
    raise EOpcClient.CreateRes(@SRequiresActiveServer)
end;

constructor TOpcSimpleClient.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  FGroups:= TOpcGroupCollection.Create(Self)
end;

destructor TOpcSimpleClient.Destroy;
begin
  FGroups.Free;
  inherited Destroy
end;

function TOpcSimpleClient.GetActive: Boolean;
begin
  if csLoading in ComponentState then
    Result:= FActive
  else
    Result:= Assigned(FOpcServer)
end;

procedure TOpcSimpleClient.Connect(NoGroups: Boolean);
var
  Clsid: TGUID;
  WideHostName: WideString;
  CPC: IConnectionPointContainer;
  CP: IConnectionPoint;
begin
  if not Assigned(FOpcServer) then
  begin
    if FProgID = '' then
      raise EOpcClient.CreateRes(@sNoProgID);
    if csDesigning in ComponentState then   {cf 1.16d.1}
      GetServerInfo(FHostName, FProgID, Clsid, FServerDesc, FVendorName)
    else
      GetServerInfo(FHostName, FProgID, Clsid, FServerDesc, FVendorName, [soUseOpcEnum]);
    if IsLocalHost(FHostName) then
      FOpcServer:= CreateComObject(Clsid) as IOPCServer
    else
      FOpcServer:= CreateRemoteComObject(FHostName, Clsid) as IOPCServer;
    if FHostName <> '' then
    begin
      WideHostName:= FHostName;
      with FOpcServer as IOPCCommon do
        SetClientName(PWideChar(WideHostName))
    end;
    FShutdownCookie:= 0;
    if (FOpcServer.QueryInterface(IConnectionPointContainer, CPC) = S_OK) and
       (CPC.FindConnectionPoint(IOpcShutdown, CP) = S_OK) then
      CP.Advise(TShutdownSink.Create(Self), FShutdownCookie);
    if not (csDesigning in ComponentState) and not NoGroups then
      ConnectGroups
  end;
  FActive:= true;
  if Assigned(FOnConnect) then
    FOnConnect(Self);
  UpdateBrowser
end;

procedure TOpcSimpleClient.ConnectGroups;
begin
  Groups.ServerAddGroups(FOpcServer)
end;

procedure TOpcSimpleClient.DisconnectGroups;
begin
  Groups.ServerRemoveGroups(FOpcServer)
end;

procedure TOpcSimpleClient.Disconnect;
var
  CPC: IConnectionPointContainer;
  CP: IConnectionPoint;
begin
  if Assigned(FOpcServer) then
  begin
    if FShutdownCookie <> 0 then
    begin
      if (FOpcServer.QueryInterface(IConnectionPointContainer, CPC) = S_OK) and
         (CPC.FindConnectionPoint(IOpcShutdown, CP) = S_OK) then
        CP.UnAdvise(FShutdownCookie);
      FShutdownCookie:= 0
    end;
    if not (csDesigning in ComponentState) then
      DisconnectGroups;
    FOpcServer:= nil;
    FServerDesc:= ''
  end;
  FActive:= false;
  if Assigned(FOnDisconnect) then
    FOnDisconnect(Self);
  UpdateBrowser
end;

procedure TOpcSimpleClient.UpdateBrowser;
var
  Msg: TMessage;
begin
  if Assigned(FBrowserWindow) then
  begin
    FillChar(Msg, SizeOf(Msg), 0);
    Msg.Msg:= WM_OPCBROWSEUPDATE;
    Msg.WParam:= Integer(Self);
    FBrowserWindow.Dispatch(Msg)
  end
end;

procedure TOpcSimpleClient.SetActive(Value: Boolean);
begin
  if Value <> FActive then
  begin
    if csLoading in ComponentState then
    begin
      FActive:= Value
    end else
    begin
      if Value then
        Connect
      else
        Disconnect
    end
  end
end;

procedure TOpcSimpleClient.SetHostName(const Value: String);
var
  WideName: WideString;
begin
  if FHostName <> Value then
  begin
    FHostName:= Value;
    if Assigned(FOpcServer) then
    begin
      WideName:= Value;
      with FOpcServer as IOPCCommon do
        SetClientName(PWideChar(WideName))
    end
  end
end;

procedure TOpcSimpleClient.SetProgID(const Value: String);
begin
  CheckNotActive;
  FProgID:= Value;
end;

procedure TOpcSimpleClient.GetBrowser(var Browse: IOPCBrowseServerAddressSpace);
begin
  CheckActive;
  if FOpcServer.QueryInterface(IOPCBrowseServerAddressSpace, Browse) <> S_OK then
    raise EOpcClient.CreateRes(@SIOPCBrowseNotSupported)
end;

procedure TOpcSimpleClient.GetAllItems(ItemIDs: TStrings);
{if you pass OPC_FLAT to a Hierarchical space you are supposed to
get a full list of all item_ids. ICONICS sample server, for one,
does not do that. Don't these people ever bother to read the spec?
The SST Sample server seems to work correctly. I expect I will end
up recursing through non-compliant servers myself. GRRR!}
var
  Browse: IOPCBrowseServerAddressSpace;
  Enum: IEnumString;
  HR: HRESULT;
  Res: PWideChar;
  Alloc: IMalloc;
  C: WideChar;
  NST: TOleEnum;

procedure BrowseBranch;
var
  Enum: IEnumString;
  HR: HResult;
  BrowseID: PWideChar;
  ItemID: PWideChar;
begin
  HR:= Browse.BrowseOpcItemIds(OPC_LEAF, @C, VT_EMPTY, 0, Enum);
  if HR <> S_FALSE then
  begin
    OleCheck(HR);
    while Enum.Next(1, BrowseID, nil) = S_OK do
    begin
      OleCheck(Browse.GetItemID(BrowseID, ItemID));
      ItemIds.Add(ItemID);
      Alloc.Free(ItemID);
      Alloc.Free(BrowseID)
    end
  end;
  HR:= Browse.BrowseOpcItemIds(OPC_BRANCH, @C, VT_EMPTY, 0, Enum);
  if HR <> S_FALSE then
  begin
    OleCheck(HR);
    while Enum.Next(1, BrowseID, nil) = S_OK do
    begin
      OleCheck(Browse.ChangeBrowsePosition(OPC_BROWSE_DOWN, BrowseID));
      Alloc.Free(BrowseID);
      BrowseBranch;
      OleCheck(Browse.ChangeBrowsePosition(OPC_BROWSE_UP, @C))
    end
  end
end;

begin
  GetBrowser(Browse);
  C:= #0;
  CoGetMalloc(1, Alloc);
  OleCheck(Browse.QueryOrganization(NST));
  ItemIds.BeginUpdate;
  try
    if NST = OPC_NS_FLAT then
    begin
      HR:= Browse.BrowseOpcItemIds(OPC_FLAT, @C, VT_EMPTY, 0, Enum);
      if HR <> S_FALSE then
      begin
        OleCheck(HR);
        while Enum.Next(1, Res, nil) = S_OK do
        begin
          ItemIds.Add(Res);
          Alloc.Free(Res)
        end
      end
    end else
    begin  {I should not have to do this!}
      HR:= Browse.ChangeBrowsePosition(OPC_BROWSE_TO, @C);
      if HR = E_INVALIDARG then
      begin
        repeat
          HR:= Browse.ChangeBrowsePosition(OPC_BROWSE_UP, @C)
        until HR <> S_OK;
      end else
      begin
        OleCheck(HR)
      end;
      BrowseBranch
    end
  finally
    ItemIds.EndUpdate
  end
end;

procedure TOpcSimpleClient.CheckNotActive;
begin
  if Active then
    raise EOpcClient.CreateRes(@SNotActiveServer)
end;

procedure TOpcSimpleClient.ServerShutdown(pzReason: PWideChar);
begin
  FShutdownReason:= pzReason;
  if Assigned(FOnServerShutdown) then
    FOnServerShutdown(Self);
  Active:= false
end;

function TOpcSimpleClient.ShutdownConnected: Boolean;
begin
  Result:= FShutdownCookie <> 0
end;

procedure TOpcSimpleClient.WMOpcClientConnect(var Msg: TMessage);
begin
  FBrowserWindow:= TObject(Msg.wParam)
end;

procedure TOpcSimpleClient.GetItemProperties(const ItemId: string;
  var Properties: TItemProperties);
var
  Props: IOpcItemProperties;
  WideId: WideString;
  dwCount: Cardinal;
  PropIDs: PDWORDARRAY;
  Descriptions: POleStrList;
  DataTypes: PVarTypeList;
  Alloc: IMalloc;
  i: Integer;
begin
  CheckActive;
  Props:= FOpcServer as IOPCItemProperties;
  WideId:= ItemId;
  Alloc:= GetMalloc;
  OpcCheck(Props.QueryAvailableProperties(PWideChar(WideId),
     dwCount, PropIDs, Descriptions, DataTypes));
  try
    SetLength(Properties, dwCount);
    for i:= 0 to dwCount - 1 do
    with Properties[i] do
    begin
      ID:= PropIDs^[i];
      Desc:= Descriptions^[i];
      Alloc.Free(Descriptions^[i]);
      Datatype:= DataTypes^[i]
    end
  finally
    Alloc.Free(PropIds);
    Alloc.Free(DataTypes);
    Alloc.Free(Descriptions)
  end;
  UpdateProperties(ItemId, Properties)
end;

procedure TOpcSimpleClient.UpdateProperties(const ItemId: string;
  var Properties: TItemProperties);
var
  Props: IOpcItemProperties;
  WideId: WideString;
  dwCount: Cardinal;
  PropIDs: array of DWORD;
  Data: POleVariantArray;
  Errors: PResultList;
  i: Integer;
begin
  CheckActive;
  Props:= FOpcServer as IOPCItemProperties;
  WideId:= ItemId;
  dwCount:= Length(Properties);
  SetLength(PropIDs, dwCount);
  for i:= 0 to dwCount - 1 do
    PropIDs[i]:= Properties[i].ID;
  OpcCheck(Props.GetItemProperties(PWideChar(WideId),
     dwCount, Pointer(PropIDs), Data, Errors));
  try
    for i:= 0 to dwCount - 1 do
    with Properties[i] do
    begin
      Value:= Data^[i];
      Error:= Errors^[i]
    end
  finally
    FreeOleVariantArray(dwCount, Data);
    FreeResultList(Errors)
  end
end;

function TOpcSimpleClient.SupportsBrowsing: Boolean;
var
  Unk: IUnknown;
begin
  Result:= Assigned(FOpcServer) and
    (FOpcServer.QueryInterface(IOPCBrowseServerAddressSpace, Unk) = S_OK)
end;

procedure TOpcSimpleClient.SetGroups(Value: TOpcGroupCollection);
begin
  FGroups.Assign(Value)
end;

procedure TOpcSimpleClient.OpcCheck(Res: HRESULT);
begin
  if not Succeeded(Res) then
    raise EOpcClient.CreateHResult(Self, Res)
end;

function TOpcSimpleClient.GetErrorString(Res: HRESULT): string;
begin
  Result:= GetOpcErrorString(FOpcServer, Res)
end;

procedure TOpcSimpleClient.SetSortedItemLists(const Value: Boolean); {cf 1.14.1}
var
  i: Integer;
begin
  if FSortedItemLists <> Value then
  begin
    FSortedItemLists:= Value;
    for i:= 0 to Groups.Count - 1 do
      Groups[i].SortedItemList:= Value
  end
end;

procedure TOpcSimpleClient.BrowseItems(BrowseNodeEvent: TBrowseNodeEvent);
var
  Browse: IOPCBrowseServerAddressSpace;
  Enum: IEnumString;
  HR: HRESULT;
  Res: PWideChar;
  Alloc: IMalloc;
  C: WideChar;
  NST: TOleEnum;
{===}
procedure BrowseBranch(Parent: Pointer);
var
  Enum: IEnumString;
  HR: HResult;
  BrowseID: PWideChar;
  ItemID: PWideChar;
begin
  HR:= Browse.BrowseOpcItemIds(OPC_LEAF, @C, VT_EMPTY, 0, Enum);
  if HR <> S_FALSE then
  begin
    OleCheck(HR);
    while Enum.Next(1, BrowseID, nil) = S_OK do
    begin
      OleCheck(Browse.GetItemID(BrowseID, ItemID));
      BrowseNodeEvent(Parent, BrowseID, ItemID);
      Alloc.Free(ItemID);
      Alloc.Free(BrowseID)
    end
  end;
  HR:= Browse.BrowseOpcItemIds(OPC_BRANCH, @C, VT_EMPTY, 0, Enum);
  if HR <> S_FALSE then
  begin
    OleCheck(HR);
    while Enum.Next(1, BrowseID, nil) = S_OK do
    begin
      OleCheck(Browse.ChangeBrowsePosition(OPC_BROWSE_DOWN, BrowseID));
      BrowseBranch(BrowseNodeEvent(Parent, BrowseID, ''));
      Alloc.Free(BrowseID);
      OleCheck(Browse.ChangeBrowsePosition(OPC_BROWSE_UP, @C))
    end
  end
end;
{===}
begin
  if Assigned(OpcServer) and
     (OpcServer.QueryInterface(IOpcBrowseServerAddressSpace, Browse) = S_OK) then
  begin
    C:= #0;
    CoGetMalloc(1, Alloc);
    OleCheck(Browse.QueryOrganization(NST));
    if NST = OPC_NS_FLAT then
    begin
      HR:= Browse.BrowseOpcItemIds(OPC_FLAT, @C, VT_EMPTY, 0, Enum);
      if HR <> S_FALSE then
      begin
        OleCheck(HR);
        while Enum.Next(1, Res, nil) = S_OK do
        begin
          BrowseNodeEvent(nil, Res, Res);
          Alloc.Free(Res)
        end
      end
    end else
    begin
      HR:= Browse.ChangeBrowsePosition(OPC_BROWSE_TO, @C);
      if HR = E_INVALIDARG then
      begin
        repeat
          HR:= Browse.ChangeBrowsePosition(OPC_BROWSE_UP, @C)
        until HR <> S_OK;
      end else
      begin
        OleCheck(HR)
      end;
      BrowseBranch(nil)
    end;
  end;
end;

function TOpcSimpleClient.GetServerStatus: TServerStatus;
var
  ppServerStatus: POPCSERVERSTATUS;
begin
  if not Assigned(OpcServer) then
    raise EOpcClient.CreateRes(@SNoOpcServer);

  FillChar(Result, SizeOf(Result), 0);

  ppServerStatus := nil;
  OpcServer.GetStatus(ppServerStatus);

  Result.StartTime := FileTimeToDateTime(ppServerStatus.ftStartTime);
  Result.CurrentTime := FileTimeToDateTime(ppServerStatus.ftCurrentTime);
  Result.LastUpdateTime := FileTimeToDateTime(ppServerStatus.ftLastUpdateTime);
  Result.ServerState := ppServerStatus.dwServerState;
  Result.GroupCount := ppServerStatus.dwGroupCount;
  Result.BandWidth := ppServerStatus.dwBandWidth;
  Result.MajorVersion := ppServerStatus.wMajorVersion;
  Result.MinorVersion := ppServerStatus.wMinorVersion;
  Result.BuildNumber := ppServerStatus.wBuildNumber;
  Result.Reserved := ppServerStatus.wReserved;
  Result.VendorInfo := ppServerStatus.szVendorInfo;

  FreeServerStatus(ppServerStatus);
end;

{ TDa1Sink }

constructor TDa1Sink.Create(aGroup: TOpcGroup);
begin
  inherited Create;
  FGroup:= aGroup
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
      FGroup.DA1DataTimeChange(GroupHeader^.dwItemCount, DataTimeItemHeaders, Data)
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

constructor TDa2Sink.Create(aGroup: TOpcGroup);
begin
  inherited Create;
  FGroup:= aGroup
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
  if hGroup = DWORD(FGroup) then
    FGroup.DataChange(dwCount, phClientItems, pvValues,
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
  if hGroup = DWORD(FGroup) then
    FGroup.WriteComplete(dwCount, pClienthandles, pErrors);
  Result:= S_OK
end;

{ TShutdownSink }

constructor TShutdownSink.Create(aClient: TOpcSimpleClient);
begin
  inherited Create;
  FClient:= aClient
end;

function TShutdownSink.ShutdownRequest(szReason: POleStr): HResult;
begin
  FClient.ServerShutdown(szReason);
  Result:= S_OK
end;

{ EOpcClient }

constructor EOpcClient.CreateHResult(aClient: TOpcSimpleClient; Res: HRESULT);
begin
  if Assigned(aClient) then
    Create(aClient.GetErrorString(Res))
  else
    Create(GetOpcErrorString(nil, Res))
end;

initialization
  CoInitialize(nil);
  InitialiseClientSecurity
finalization
  CoUninitialize
end.

