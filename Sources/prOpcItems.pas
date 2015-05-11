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
unit prOpcItems;

interface

uses Windows, SysUtils, ActiveX, Variants,
     prOpcError, prOpcDa, prOpcTypes, prOpcClasses, prOpcServer;

type
  TVQT = record
    Value : Variant;
    Quality : Word;
    Timestamp : TFileTime;
  end;

  TOPCDataItem = class;
  TOPCWrite = function (Sender : TOPCDataItem; Value : OleVariant) : Boolean of object;
  TOPCSetQuality = function (Sender : TOPCDataItem; Quality : LongWord) : Boolean of object;
  TOPCSetTimestamp = function (Sender : TOPCDataItem; Timestamp : TFileTime) : Boolean of object;
  TOPCDataItem = class(TObject)
  private
    FValue: OleVariant;
    FQuality : Word;
    FTimestamp: TFileTime;
    FOnWrite: TOPCWrite;
    FOnSetQuality: TOPCSetQuality;
    FOnSetTimestamp: TOPCSetTimestamp;
    FQualityGood: Boolean;
    DestroyScheduled: Boolean;
    procedure SetValue(const Value: OleVariant);
    procedure SetQualityGood(const Value: Boolean);
    procedure SetQuality(const Value: Word);
    procedure InternalSetVQT(aValue: Variant; aQuality: Word; aTimeStamp: TFileTime);
    function GetVQT: TVQT;
    procedure SetVQT(const Value: TVQT);
    procedure UpdateValue;
  protected
  public
    Descr : string[255];
    Units : string[255];
    FVarType: Integer;
    UpdateEvent : TSubscriptionEvent;
    EUInfo: IEUInfo;
    ConnectionCount : Integer;

    constructor Create(VarType : Integer; aQualityGood : Boolean = True; aUnits : string = ''; aDescr : string = '');
    procedure SafeDestroy;
    function AccessRights: TAccessRights;
    function GetItemProperties: TItemProperties;

    property Value: OleVariant read FValue write SetValue;
    property QualityGood : Boolean read FQualityGood write SetQualityGood;
    property Quality : Word read FQuality write SetQuality;
    property Timestamp : TFileTime read FTimestamp;
    property VQT : TVQT read GetVQT write SetVQT;
    property OnWrite: TOPCWrite read FOnWrite write FOnWrite;
    property OnSetQuality: TOPCSetQuality read FOnSetQuality write FOnSetQuality;
    property OnSetTimestamp: TOPCSetTimestamp read FOnSetTimestamp write FOnSetTimestamp;
  end;

  TItemPropertyStatic = class(TItemProperty)
  private
    FPropDescription: string;
    FPropValue: OleVariant;
  public
    function Description: string; override;
    function DataType: Integer; override;
    function GetPropertyValue: OleVariant; override;
    constructor Create(aPid: Integer; PropDescription: string; PropValue: OleVariant);
  end;

  TItemPropertyDescription = class(TItemPropertyStatic)
  public
    constructor Create(Description : string);
  end;

  TItemPropertyUnits = class(TItemPropertyStatic)
  public
    constructor Create(Units : string);
  end;

  TOPCDataItemServer = class(TOpcItemServer)
  protected
    function Options: TServerOptions; override;
    function CheckItemHandle(ItemHandle: TItemHandle) : Boolean;
    function SubscribeToItem(ItemHandle: TItemHandle; UpdateEvent: TSubscriptionEvent): Boolean; override;
    procedure UnsubscribeToItem(ItemHandle: TItemHandle); override;
    procedure OnAddItem(Item: TGroupItemInfo); override;
    procedure OnRemoveItem(Item: TGroupItemInfo); override;
    function GetExtendedItemInfo(const ItemID: String; var AccessPath: String;
      var AccessRights: TAccessRights; var EUInfo: IEUInfo;
      var ItemProperties: IItemProperties): Integer; override;
    function GetItemVQT(ItemHandle: TItemHandle; var Quality: Word;
      var Timestamp: TFileTime): OleVariant; override;
    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); override;
    procedure SetItemQuality(ItemHandle: TItemHandle; const Quality: Word); override;
    procedure SetItemTimestamp(ItemHandle: TItemHandle; const Timestamp: TFileTime); override;
    procedure ItemDestroyed; virtual;
    function OnItemNotFound(const ItemID: String; var AccessPath: String) : TNamespaceNode; virtual;
  public
    constructor Create;
    destructor Destroy; override;
  end;

  function AddOPCItem(ItemIDList: TItemIDList; Name: string; OPCDataItem: TOPCDataItem) : TNamespaceNode;
  function CountOPCItems(ItemIDList: TItemIDList) : Integer;

implementation

var
  OPCDataItemServer : TOPCDataItemServer;

function AddOPCItem(ItemIDList: TItemIDList; Name: string; OPCDataItem: TOPCDataItem) : TNamespaceNode;
begin
  try
    Result := ItemIDList.AddItemID(Name, OPCDataItem.AccessRights, OPCDataItem.FVarType);
    Result.Data := OPCDataItem;
  except
   on E : Exception do
    raise Exception.Create(ItemIDList.Path + ':' + E.Message);
  end;
end;

function CountOPCItems(ItemIDList: TItemIDList) : Integer;
var
  I : Integer;
begin
  Result := 0;
  for I := 0 to ItemIDList.ChildCount-1 do
   begin
     if ItemIDList.Child(I) is TNamespaceItem then
      Inc(Result);
     if ItemIDList.Child(I) is TItemIDList then
      Result := Result + CountOPCItems(TItemIDList(ItemIDList.Child(I)));
   end;
end;

{ TOPCDataItem }

function TOPCDataItem.AccessRights: TAccessRights;
begin
  Result := [iaRead];
  if Assigned(FOnWrite) then
   Result := Result + [iaWrite];
end;

function TOPCDataItem.GetItemProperties: TItemProperties;
var
  Limits : OleVariant;
begin
  Result := TItemProperties.Create;
  Result.Add(TItemPropertyUnits.Create(Units));
  Result.Add(TItemPropertyDescription.Create(Descr));
  if (EUInfo <> nil) and
     (EUInfo.EUType = euAnalog) then
   begin
     Limits := EUInfo.EUInfo;
     Result.Add(TItemPropertyStatic.Create(OPC_PROPERTY_HIGH_EU, OPC_PROPERTY_DESC_HIGH_EU, Limits[0]));
     Result.Add(TItemPropertyStatic.Create(OPC_PROPERTY_LOW_EU,  OPC_PROPERTY_DESC_LOW_EU,  Limits[1]));
   end;
end;

constructor TOPCDataItem.Create(VarType: Integer; aQualityGood : Boolean = True; aUnits : string = ''; aDescr : string = '');
begin
  inherited Create;
  QualityGood := aQualityGood;
  GetSystemTimeAsFileTime(FTimestamp);
  Self.FVarType := VarType;
  case VarType of
    varOleStr : VariantChangeType(FValue, '', 0, VarType);
    else VariantChangeType(FValue, 0, 0, VarType);
  end;

  Units := aUnits;
  Descr := aDescr;
end;

procedure TOPCDataItem.SafeDestroy;
begin
  OnWrite := nil;
  OnSetQuality := nil;
  OnSetTimestamp := nil;
  FQuality := 0;
  UpdateValue;
  if not Assigned(UpdateEvent) then
   Free
  else
   DestroyScheduled := True;

  if Assigned(OPCDataItemServer) then
   OPCDataItemServer.ItemDestroyed;
end;

procedure TOPCDataItem.SetQualityGood(const Value: Boolean);
begin
  if FQualityGood <> Value then
   begin
     FQualityGood := Value;

     if Value then
      InternalSetVQT(FValue, OPC_QUALITY_GOOD, TimestampNotSet)
     else
      InternalSetVQT(FValue, OPC_QUALITY_BAD or OPC_QUALITY_COMM_FAILURE, TimestampNotSet);
   end;
end;

procedure TOPCDataItem.SetQuality(const Value: Word);
begin
  InternalSetVQT(FValue, Value, TimestampNotSet);
end;

function TOPCDataItem.GetVQT: TVQT;
begin
  Result.Value := Value;
  Result.Quality := Quality;
  Result.Timestamp := Timestamp;
end;

procedure TOPCDataItem.SetVQT(const Value: TVQT);
begin
  InternalSetVQT(Value.Value, Value.Quality, Value.Timestamp);
end;

procedure TOPCDataItem.SetValue(const Value: OleVariant);
begin
  InternalSetVQT(Value, FQuality, TimestampNotSet);
end;

procedure TOPCDataItem.InternalSetVQT(aValue: Variant; aQuality: Word;
  aTimeStamp: TFileTime);
var
  NeedUpdate : Boolean;
begin
  NeedUpdate := False;

  if FValue <> aValue then
   begin
     VariantChangeType(FValue, aValue, 0, FVarType);
     NeedUpdate := True;
   end;

  if FQuality <> aQuality then
   begin
     FQuality := aQuality;

     FQualityGood := FQuality = OPC_QUALITY_GOOD;

     NeedUpdate := True;
   end;

  if NeedUpdate then
   begin
     if Int64(aTimeStamp) = Int64(TimestampNotSet) then
      GetSystemTimeAsFileTime(FTimestamp)
     else
      FTimestamp := aTimestamp;

     UpdateValue;
   end;
end;

procedure TOPCDataItem.UpdateValue;
begin
  if Assigned(UpdateEvent) then
   UpdateEvent(FValue, FQuality, FTimestamp);
end;

{ TItemPropertyStatic }

constructor TItemPropertyStatic.Create(aPid: Integer; PropDescription: string; PropValue: OleVariant);
begin
  inherited Create(aPid);
  FPropDescription := PropDescription;
  FPropValue := PropValue
end;

function TItemPropertyStatic.DataType: Integer;
begin
  Result := VarType(FPropValue);
end;

function TItemPropertyStatic.Description: string;
begin
  Result := FPropDescription
end;

function TItemPropertyStatic.GetPropertyValue: OleVariant;
begin
  Result := FPropValue;
end;

{ TItemPropertyDescription }

constructor TItemPropertyDescription.Create(Description: string);
begin
  inherited Create(OPC_PROPERTY_DESCRIPTION, OPC_PROPERTY_DESC_DESCRIPTION, Description);
end;

{ TItemPropertyUnits }

constructor TItemPropertyUnits.Create(Units: string);
begin
  inherited Create(OPC_PROPERTY_EU_UNITS, OPC_PROPERTY_DESC_EU_UNITS, Units);
end;

{ TOPCDataItemServer }

constructor TOPCDataItemServer.Create;
begin
  inherited Create;
  OPCDataItemServer := Self;
end;

destructor TOPCDataItemServer.Destroy;
begin
  OPCDataItemServer := nil;
  inherited;
end;

function TOPCDataItemServer.Options: TServerOptions;
begin
  Result := inherited Options + [soHierarchicalBrowsing];
end;

function TOPCDataItemServer.CheckItemHandle(
  ItemHandle: TItemHandle): Boolean;
begin
  if (ItemHandle <> 0) and
     TObject(ItemHandle).InheritsFrom(TOPCDataItem) then
   Result := True
  else
   Result := False;
end;

function TOPCDataItemServer.SubscribeToItem(ItemHandle: TItemHandle;
  UpdateEvent: TSubscriptionEvent): Boolean;
begin
  if CheckItemHandle(ItemHandle) then
   begin
     TOPCDataItem(ItemHandle).UpdateEvent := UpdateEvent;
     Result := True;
   end
  else
   Result := False;
end;

procedure TOPCDataItemServer.UnsubscribeToItem(ItemHandle: TItemHandle);
begin
  if CheckItemHandle(ItemHandle) then
   begin
     with TOPCDataItem(ItemHandle) do
      begin
        UpdateEvent := nil;
        if DestroyScheduled then
         Free;
      end;
   end;
end;

procedure TOPCDataItemServer.OnAddItem(Item: TGroupItemInfo);
begin
  if CheckItemHandle(Item.ItemHandle) then
   Inc(TOPCDataItem(Item.ItemHandle).ConnectionCount);
end;

procedure TOPCDataItemServer.OnRemoveItem(Item: TGroupItemInfo);
begin
  if CheckItemHandle(Item.ItemHandle) then
   Dec(TOPCDataItem(Item.ItemHandle).ConnectionCount);
end;

function TOPCDataItemServer.GetExtendedItemInfo(const ItemID: String;
  var AccessPath: String; var AccessRights: TAccessRights;
  var EUInfo: IEUInfo; var ItemProperties: IItemProperties): Integer;
var
  Node : TNamespaceNode;
  OPCItem : TOPCDataItem;
begin
  Node := RootNode.Find(ItemID);

  if Node = nil then
   Node := OnItemNotFound(ItemID, AccessPath);

  if Node <> nil then
   begin
     if not CheckItemHandle(Integer(Node.Data)) then
      raise EOpcError.Create(OPC_E_INVALIDITEMID);

     OPCItem := Node.Data;
     Result := Integer(Node.Data);
     AccessRights := OPCItem.AccessRights;
     EUInfo := OPCItem.EUInfo;
     ItemProperties:= OPCItem.GetItemProperties
   end
  else
   raise EOpcError.Create(OPC_E_INVALIDITEMID)
end;

function TOPCDataItemServer.GetItemVQT(ItemHandle: TItemHandle;
  var Quality: Word; var Timestamp: TFileTime): OleVariant;
begin
 if CheckItemHandle(ItemHandle) then
   begin
     Result := TOPCDataItem(ItemHandle).Value;
     Quality := TOPCDataItem(ItemHandle).Quality;
     Timestamp := TOPCDataItem(ItemHandle).Timestamp;
   end
   else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
end;

procedure TOPCDataItemServer.SetItemValue(ItemHandle: TItemHandle;
  const Value: OleVariant);
begin
  if CheckItemHandle(ItemHandle) then
   if Assigned(TOPCDataItem(ItemHandle).OnWrite) then
    begin
      try
        if not TOPCDataItem(ItemHandle).OnWrite(TOPCDataItem(ItemHandle), Value) then
         raise EOpcError.Create(OPC_E_RANGE)
      except
        raise EOpcError.Create(OPC_E_RANGE)
      end
    end
   else
    EOpcError.Create(OPC_E_BADRIGHTS)
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
end;

procedure TOPCDataItemServer.SetItemQuality(ItemHandle: TItemHandle;
  const Quality: Word);
begin
  if CheckItemHandle(ItemHandle) then
   if Assigned(TOPCDataItem(ItemHandle).OnSetQuality) then
    TOPCDataItem(ItemHandle).OnSetQuality(TOPCDataItem(ItemHandle), Quality)
  else
    TOPCDataItem(ItemHandle).FQuality := Quality
end;

procedure TOPCDataItemServer.SetItemTimestamp(ItemHandle: TItemHandle;
  const Timestamp: TFileTime);
begin
  if CheckItemHandle(ItemHandle) then
   if Assigned(TOPCDataItem(ItemHandle).OnSetTimestamp) then
    TOPCDataItem(ItemHandle).OnSetTimestamp(TOPCDataItem(ItemHandle), Timestamp)
  else
    TOPCDataItem(ItemHandle).FTimestamp := Timestamp
end;

procedure TOPCDataItemServer.ItemDestroyed;
begin
end;

function TOPCDataItemServer.OnItemNotFound(const ItemID: String; var AccessPath: String) : TNamespaceNode;
begin
  Result := nil;
end;

end.
