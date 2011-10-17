unit OpcServerUnit;

interface

uses
  SysUtils, Classes, prOpcServer, prOpcTypes;

type
  TDemo10 = class(TOpcItemServer)
  private
  protected
    function Options: TServerOptions; override;
  public
    SmallintData: array[0..4] of Smallint;
    LongintData: array[0..4] of Longint;
    SingleData: array[0..4] of Single;
    DoubleData: array[0..4] of Double;
    TDateTimeData: array[0..4] of TDateTime;
    WideStringData: array[0..4] of WideString;
    BooleanData: array[0..4] of Boolean;
    ByteData: array[0..4] of Byte;
    function GetItemInfo(const ItemID: String; var AccessPath: string;
      var AccessRights: TAccessRights): Integer; override;
    procedure ReleaseHandle(ItemHandle: TItemHandle); override;
    procedure ListItemIds(List: TItemIDList); override;
    function GetItemValue(ItemHandle: TItemHandle;
                            var Quality: Word): OleVariant; override;
    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); override;
    constructor Create;
  end;

implementation
uses
  prOpcError, prOpcVarUtils;

{ TDemo10 }
const
  SmallintDataHandle = 1;
  LongintDataHandle = 2;
  SingleDataHandle = 3;
  DoubleDataHandle = 4;
  TDateTimeDataHandle = 5;
  WideStringDataHandle = 6;
  BooleanDataHandle = 7;
  ByteDataHandle = 8;


function TDemo10.Options: TServerOptions;
begin
  Result:= [soAlwaysAllocateErrorArrays]
end;

function TDemo10.GetItemInfo(const ItemID: String; var AccessPath: string;
  var AccessRights: TAccessRights): Integer;
begin
  {Return a handle that will subsequently identify ItemID}
  {raise exception of type EOpcError if Item ID not recognised}
  if SameText(ItemId, 'SmallintArrayItem') then
  begin
    Result:= SmallintDataHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  if SameText(ItemId, 'LongintArrayItem') then
  begin
    Result:= LongintDataHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  if SameText(ItemId, 'SingleArrayItem') then
  begin
    Result:= SingleDataHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  if SameText(ItemId, 'DoubleArrayItem') then
  begin
    Result:= DoubleDataHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  if SameText(ItemId, 'TDateTimeArrayItem') then
  begin
    Result:= TDateTimeDataHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  if SameText(ItemId, 'WideStringArrayItem') then
  begin
    Result:= WideStringDataHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  if SameText(ItemId, 'BooleanArrayItem') then
  begin
    Result:= BooleanDataHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  if SameText(ItemId, 'ByteArrayItem') then
  begin
    Result:= ByteDataHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  begin
    raise EOpcError.Create(OPC_E_INVALIDITEMID)
  end
end;

procedure TDemo10.ReleaseHandle(ItemHandle: TItemHandle);
begin
  {Release the handle previously returned by GetItemInfo}
  {DO NOT DELETE THIS EMPTY IMPLEMENTATION}
end;

procedure TDemo10.ListItemIds(List: TItemIDList);
begin
  {Call List.AddItemId(ItemId, AccessRights, VarType) for each ItemId}
  List.AddItemId('SmallintArrayItem', [iaRead, iaWrite], varSmallint and varArray);
  List.AddItemId('LongintArrayItem', [iaRead, iaWrite], varInteger and varArray);
  List.AddItemId('SingleArrayItem', [iaRead, iaWrite], varSingle and varArray);
  List.AddItemId('DoubleArrayItem', [iaRead, iaWrite], varDouble and varArray);
  List.AddItemId('TDateTimeArrayItem', [iaRead, iaWrite], varDate and varArray);
  List.AddItemId('WideStringArrayItem', [iaRead, iaWrite], varOleStr and varArray);
  List.AddItemId('BooleanArrayItem', [iaRead, iaWrite], varBoolean and varArray);
  List.AddItemId('ByteArrayItem', [iaRead, iaWrite], varByte and varArray);
end;

function TDemo10.GetItemValue(ItemHandle: TItemHandle;
                           var Quality: Word): OleVariant;
begin
  {return the value of the item identified by ItemHandle}
  case ItemHandle of
    SmallintDataHandle: Result:= CreateVarArray(SmallintData);
    LongintDataHandle: Result:= CreateVarArray(LongintData);
    SingleDataHandle: Result:= CreateVarArray(SingleData);
    DoubleDataHandle: Result:= CreateVarArray(DoubleData);
    TDateTimeDataHandle: Result:= CreateVarArray(TDateTimeData);
    WideStringDataHandle: Result:= CreateVarArray(WideStringData);
    BooleanDataHandle: Result:= CreateVarArray(BooleanData);
    ByteDataHandle: Result:= CreateVarArray(ByteData);
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
  end
end;

procedure TDemo10.SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant);
begin
  {set the value of the item identified by ItemHandle}
  case ItemHandle of
    SmallintDataHandle: ReadVarArray(Value, SmallintData);
    LongintDataHandle: ReadVarArray(Value, LongintData);
    SingleDataHandle: ReadVarArray(Value, SingleData);
    DoubleDataHandle: ReadVarArray(Value, DoubleData);
    TDateTimeDataHandle: ReadVarArray(Value, TDateTimeData);
    WideStringDataHandle: ReadVarArray(Value, WideStringData);
    BooleanDataHandle: ReadVarArray(Value, BooleanData);
    ByteDataHandle: ReadVarArray(Value, ByteData);
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
  end
end;

const
  ServerGuid: TGUID = '{C095C703-77BA-42F1-B6AA-6C234C25496E}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Array items no Rtti';
  ServerVendor = 'Production Robots Eng. Ltd';

constructor TDemo10.Create;
begin
  inherited Create;
  WideStringData[0]:= 'a';
  WideStringData[1]:= 'b';
  WideStringData[2]:= 'c';
  WideStringData[3]:= 'd';
  WideStringData[4]:= 'e';
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo10.Create)
end.
