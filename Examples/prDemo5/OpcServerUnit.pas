{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit OpcServerUnit;

interface

uses
  SysUtils, Classes, prOpcRttiServer, prOpcServer, prOpcTypes;

type
  TDemo5 = class(TRttiItemServer)
  private
    procedure LogMsg(const Msg: string);
    function GetFormHeight: Integer;
    function GetFormWidth: Integer;
    function GetTickCount: Integer;
    function GetTimeOfDay: String;
    procedure SetFormHeight(Value: Integer);
    procedure SetFormWidth(Value: Integer);
  protected
    procedure OnClientConnect(Client: TClientInfo); override;
    procedure OnClientDisconnect(Client: TClientInfo); override;
    procedure OnClientSetName(Client: TClientInfo); override;
    procedure OnAddGroup(Group: TGroupInfo); override;
    procedure OnRemoveGroup(Group: TGroupInfo); override;
    procedure OnAddItem(Item: TGroupItemInfo); override;
    procedure OnRemoveItem(Item: TGroupItemInfo); override;
  public
  published
    property TickCount: Integer read GetTickCount;
    property TimeOfDay: String read GetTimeOfDay;
    property FormWidth: Integer read GetFormWidth write SetFormWidth;
    property FormHeight: Integer read GetFormHeight write SetFormHeight;
  end;

implementation
uses
  prOpcError, Windows, MainUnit, ComCtrls, TypInfo;

{ TDemo5 }

procedure TDemo5.OnClientConnect(Client: TClientInfo);
var
  NewNode: TTreeNode;
begin
  {Code here will execute whenever a client connects}
  NewNode:= MainForm.TreeView.Items.Add(nil,
     Format('New Client started at %s',
       [DateTimeToStr(Now)]));
  NewNode.Data:= Client;
  Client.Data:= NewNode;
  MainForm.ExitServerButton.Enabled:= false;
  LogMsg('Client connected Ok')
end;

procedure TDemo5.OnClientDisconnect(Client: TClientInfo);
begin
  {Code here will execute whenever a client disconnects}
  MainForm.TreeView.Items.Delete(TTreeNode(Client.Data));
  if ClientCount = 0 then
    MainForm.ExitServerButton.Enabled:= true;
  LogMsg('Client disconnected Ok');
end;

procedure TDemo5.OnClientSetName(Client: TClientInfo);
begin
  {Code here will execute whenever a client calls IOpcCommon.SetClientName}
  TTreeNode(Client.Data).Text:=
    Format('Client: %s', [Client.ClientName]);
  LogMsg('Client SetName Ok')
end;

procedure TDemo5.OnAddGroup(Group: TGroupInfo);
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

procedure TDemo5.OnRemoveGroup(Group: TGroupInfo);
begin
  {Code here will execute whenever a client removes a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Group.Data));
  LogMsg(Format('Remove Group %s ok', [Group.Name]))
end;

procedure TDemo5.OnAddItem(Item: TGroupItemInfo);
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

procedure TDemo5.OnRemoveItem(Item: TGroupItemInfo);
begin
  {Code here will execute whenever a client removes an item from a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Item.Data));
  LogMsg(Format('Remove Item %s ok', [Item.ItemID]))
end;

procedure TDemo5.LogMsg(const Msg: string);
begin
  if Assigned(MainForm) then
    MainForm.DebugLog.Items.Add(Msg)
end;

const
  ServerGuid: TGUID = '{B070D601-12CF-11D5-944D-00C0F023FA1C}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Rtti demo';
  ServerVendor = 'Production Robots Eng. Ltd.';

function TDemo5.GetFormHeight: Integer;
begin
  Result:= MainForm.Height
end;

function TDemo5.GetFormWidth: Integer;
begin
  Result:= MainForm.Width
end;

function TDemo5.GetTickCount: Integer;
begin
  Result:= Windows.GetTickCount
end;

function TDemo5.GetTimeOfDay: String;
begin
  Result:= DateTimeToStr(Now)
end;

procedure TDemo5.SetFormHeight(Value: Integer);
begin
  MainForm.Height:= Value
end;

procedure TDemo5.SetFormWidth(Value: Integer);
begin
  MainForm.Width:= Value
end;


initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo5.Create)
end.
