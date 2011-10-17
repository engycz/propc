{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit OpcServerUnit;

interface

uses
  SysUtils, Classes, prOpcServer, prOpcTypes;

type
  TDemo4 = class(TOpcItemServer)
  private
    procedure LogMsg(const Msg: string);
  protected
    procedure OnClientConnect(Client: TClientInfo); override;
    procedure OnClientDisconnect(Client: TClientInfo); override;
    procedure OnClientSetName(Client: TClientInfo); override;
    procedure OnAddGroup(Group: TGroupInfo); override;
    procedure OnRemoveGroup(Group: TGroupInfo); override;
    procedure OnAddItem(Item: TGroupItemInfo); override;
    procedure OnRemoveItem(Item: TGroupItemInfo); override;
  public
    function GetItemInfo(const ItemID: String; var AccessPath: String;
       var AccessRights: TAccessRights): Integer; override;
    procedure ReleaseHandle(ItemHandle: TItemHandle); override;
    procedure ListItemIds(List: TItemIDList); override;
    function GetItemValue(ItemHandle: TItemHandle;
                            var Quality: Word): OleVariant; override;
    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); override;
  end;

implementation
uses
  prOpcError, Windows, MainUnit, ComCtrls;

const
  TickCountHandle = 1;
  TimeOfDayHandle = 2;
  FormWidthHandle = 3;
  FormHeightHandle = 4;

{ TDemo4 }

procedure TDemo4.OnClientConnect(Client: TClientInfo);
var
  NewNode: TTreeNode;
begin
  {Code here will execute whenever a client connects}
  NewNode:= MainForm.TreeView.Items.Add(nil,
     Format('New Client started at %s',
       [DateTimeToStr(Now)]));
  NewNode.Data:= Client;
  Client.Data:= NewNode;
  LogMsg('Client connected Ok')
end;

procedure TDemo4.OnClientDisconnect(Client: TClientInfo);
begin
  {Code here will execute whenever a client disconnects}
  MainForm.TreeView.Items.Delete(TTreeNode(Client.Data));
  LogMsg('Client disconnected Ok')
end;

procedure TDemo4.OnClientSetName(Client: TClientInfo);
begin
  {Code here will execute whenever a client calls IOpcCommon.SetClientName}
  TTreeNode(Client.Data).Text:=
    Format('Client %s',
      [Client.ClientName]);
  LogMsg('Client SetName Ok')
end;

procedure TDemo4.OnAddGroup(Group: TGroupInfo);
var
  NewNode: TTreeNode;
begin
  {Code here will execute whenever a client adds a group}
  NewNode:= MainForm.TreeView.Items.AddChild(
    TTreeNode(Group.ClientInfo.Data), Group.Name);
  NewNode.Data:= Group;
  Group.Data:= NewNode;
  LogMsg(Format('Add Group %s ok', [Group.Name]))
end;

procedure TDemo4.OnRemoveGroup(Group: TGroupInfo);
begin
  {Code here will execute whenever a client removes a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Group.Data));
  LogMsg(Format('Remove Group %s ok', [Group.Name]))
end;

procedure TDemo4.OnAddItem(Item: TGroupItemInfo);
var
  NewNode: TTreeNode;
begin
  {Code here will execute whenever a client adds an item to a group}
  NewNode:= MainForm.TreeView.Items.AddChild(
    TTreeNode(Item.Group.Data), Item.ItemID);
  NewNode.Data:= Item;
  Item.Data:= NewNode;
  LogMsg(Format('Add Item %s ok', [Item.ItemID]))
end;

procedure TDemo4.OnRemoveItem(Item: TGroupItemInfo);
begin
  {Code here will execute whenever a client removes an item from a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Item.Data));
  LogMsg(Format('Remove Item %s ok', [Item.ItemID]))
end;

function TDemo4.GetItemInfo(const ItemID: String; var AccessPath: String;
       var AccessRights: TAccessRights): Integer;
begin
  {Return a handle that will subsequently identify ItemID}
  {raise exception of type EOpcError if Item ID not recognised}
  if SameText(ItemId, 'TickCount') then
  begin
    Result:= TickCountHandle;
    AccessRights:= [iaRead]
  end else
  if SameText(ItemId, 'TimeOfDay') then
  begin
    Result:= TimeOfDayHandle;
    AccessRights:= [iaRead]
  end else
  if SameText(ItemId, 'FormWidth') then
  begin
    Result:= FormWidthHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  if SameText(ItemId, 'FormHeight') then
  begin
    Result:= FormHeightHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  begin
    raise EOpcError.Create(OPC_E_INVALIDITEMID)
  end
end;

procedure TDemo4.ReleaseHandle(ItemHandle: TItemHandle);
begin
  {Release the handle previously returned by GetItemInfo}
end;

procedure TDemo4.ListItemIds(List: TItemIDList);
begin
  {Call List.AddItemId(ItemId, AccessRights, VarType) for each ItemId}
  List.AddItemId('TickCount', [iaRead], varInteger);
  List.AddItemId('TimeOfDay', [iaRead], varOleStr);
  List.AddItemId('FormWidth', [iaRead, iaWrite], varInteger);
  List.AddItemId('FormHeight', [iaRead, iaWrite], varInteger);
end;

function TDemo4.GetItemValue(ItemHandle: TItemHandle;
                           var Quality: Word): OleVariant;
begin
  {return the value of the item identified by ItemHandle}
  case ItemHandle of
    TickCountHandle: Result:= Integer(Windows.GetTickCount);
    TimeOfDayHandle: Result:= TimeToStr(Now);
    FormWidthHandle: Result:= MainForm.Width;
    FormHeightHandle: Result:= MainForm.Height;
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
  end
end;

procedure TDemo4.SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant);
begin
  {set the value of the item identified by ItemHandle}
  case ItemHandle of
    FormWidthHandle: MainForm.Width:= Value;
    FormHeightHandle: MainForm.Height:= Value;
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
  end
end;

procedure TDemo4.LogMsg(const Msg: string);
begin
  if Assigned(MainForm) then
    MainForm.DebugLog.Items.Add(Msg)
end;

const
  ServerGuid: TGUID = '{E25AA406-121D-11D5-944C-00C0F023FA1C}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Client hook demo';
  ServerVendor = 'Production Robots Eng. Ltd.';


initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo4.Create)
end.
