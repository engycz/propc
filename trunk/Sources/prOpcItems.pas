unit prOpcItems;

interface

uses SysUtils, ActiveX, Variants,
     prOpcError, prOpcDa, prOpcTypes, prOpcClasses, prOpcServer;

type
  TOPCDataItem = class;
  TOPCWrite = function (Sender : TOPCDataItem; Value : OleVariant) : Boolean of object;
  TOPCDataItem = class(TObject)
  private
    FValue: OleVariant;
    FOnWrite: TOPCWrite;
    FQualityGood: Boolean;
    FOPCQuality : Word;
    procedure SetValue(const Value: OleVariant);
    procedure SetQualityGood(const Value: Boolean);
    procedure UpdateValue;
  protected
  public
    Descr : string[255];
    Units : string[255];
    FVarType: Integer;
    UpdateEvent : TSubscriptionEvent;

    constructor Create(VarType : Integer; QualityGood_ : Boolean = True; Units_ : string = ''; Descr_ : string = '');
    procedure SafeDestroy;
    function AccessRights: TAccessRights;

    property Value: OleVariant read FValue write SetValue;
    property QualityGood : Boolean read FQualityGood write SetQualityGood;
    property OPCQuality : Word read FOPCQuality;
    property OnWrite: TOPCWrite read FOnWrite write FOnWrite;
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
    function GetExtendedItemInfo(const ItemID: String;
                    var AccessPath: String;
                    var AccessRights: TAccessRights;
                    var EUInfo: IEUInfo;
                    var ItemProperties: IItemProperties): Integer; override;
    function GetItemValue(ItemHandle: TItemHandle;
                          var Quality: Word): OleVariant; override;
    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); override;
  end;

  function AddOPCItem(ItemIDList: TItemIDList; Name: string; OPCDataItem: TOPCDataItem) : TNamespaceNode;
  function CountOPCItems(ItemIDList: TItemIDList) : Integer;

implementation

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

constructor TOPCDataItem.Create(VarType: Integer; QualityGood_ : Boolean = True; Units_ : string = ''; Descr_ : string = '');
begin
  inherited Create;
  QualityGood := QualityGood_;
  Self.FVarType := VarType;
  case VarType of
    varOleStr : VariantChangeType(FValue, '', 0, VarType);
    else VariantChangeType(FValue, 0, 0, VarType);
  end;

  Units := Units_;
  Descr := Descr_;
end;

procedure TOPCDataItem.SafeDestroy;
begin
  Free;
end;

procedure TOPCDataItem.SetQualityGood(const Value: Boolean);
begin
  if FQualityGood <> Value then
   begin
     FQualityGood := Value;

     if Value then
      FOPCQuality := OPC_QUALITY_GOOD
     else
      FOPCQuality := OPC_QUALITY_BAD or OPC_QUALITY_COMM_FAILURE;

     UpdateValue;
   end;
end;

procedure TOPCDataItem.SetValue(const Value: OleVariant);
begin
  if FValue <> Value then
   begin
     VariantChangeType(FValue, Value, 0, FVarType);
     UpdateValue;
   end;
end;

procedure TOPCDataItem.UpdateValue;
begin
  if Assigned(UpdateEvent) then
   begin
     if QualityGood then
      UpdateEvent(FValue, OPC_QUALITY_GOOD)
     else
      UpdateEvent(FValue, OPC_QUALITY_BAD or OPC_QUALITY_COMM_FAILURE);
   end;

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
  inherited Create(OPC_PROP_DESC, 'Item Description', Description);
end;

{ TItemPropertyUnits }

constructor TItemPropertyUnits.Create(Units: string);
begin
  inherited Create(OPC_PROP_UNIT, 'Item Units', Units);
end;

{ TOPCDataItemServer }

function TOPCDataItemServer.Options: TServerOptions;
begin
  Result := inherited Options + [soHierarchicalBrowsing];
end;

function TOPCDataItemServer.CheckItemHandle(
  ItemHandle: TItemHandle): Boolean;
begin
  if (ItemHandle <> 0) and
     (TObject(ItemHandle).ClassType = TOPCDataItem) then
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
     TOPCDataItem(ItemHandle).UpdateEvent := nil;//!!CS
   end;
end;

function TOPCDataItemServer.GetExtendedItemInfo(const ItemID: String;
  var AccessPath: String; var AccessRights: TAccessRights;
  var EUInfo: IEUInfo; var ItemProperties: IItemProperties): Integer;
var
  Node : TNamespaceNode;
  OPCItem : TOPCDataItem;
begin
  Node := RootNode.Find(ItemID);

  if Node <> nil then
   begin
     if not CheckItemHandle(Integer(Node.Data)) then
      raise EOpcError.Create(OPC_E_INVALIDITEMID);

     OPCItem := Node.Data;
     Result := Integer(Node.Data);
     AccessRights := OPCItem.AccessRights;
     ItemProperties := TItemProperties.Create;
     ItemProperties.Add(TItemPropertyDescription.Create(OPCItem.Descr));
     ItemProperties.Add(TItemPropertyUnits.Create(OPCItem.Units));
   end
  else
   raise EOpcError.Create(OPC_E_INVALIDITEMID)
end;

function TOPCDataItemServer.GetItemValue(ItemHandle: TItemHandle;
  var Quality: Word): OleVariant;
begin
//!! Deleted
  if CheckItemHandle(ItemHandle) then
   begin
     Result := TOPCDataItem(ItemHandle).Value;
     Quality := TOPCDataItem(ItemHandle).OPCQuality;
   end
   else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
end;

procedure TOPCDataItemServer.SetItemValue(ItemHandle: TItemHandle;
  const Value: OleVariant);
begin
  if CheckItemHandle(ItemHandle) then
   if Assigned(TOPCDataItem(ItemHandle).OnWrite) then
    TOPCDataItem(ItemHandle).OnWrite(TOPCDataItem(ItemHandle), Value)
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
end;

end.
