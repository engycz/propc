{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{                                                            }
{ mailto engycz@gmail.com                                    }
{ http://code.google.com/p/propc/                            }
{ original source mailto: prOpcKit@prel.co.uk                }
{ original source http://www.prel.co.uk                      }
{------------------------------------------------------------}
unit OpcServerUnit;

interface

uses
  SysUtils, Classes, prOpcRttiServer, prOpcServer, prOpcTypes;

type
  TDemo8 = class(TRttiItemServer)
  private
    procedure LogMsg(const Msg: string);
  protected
    function IsRecursive: Boolean; override;
    procedure OnClientConnect(Client: TClientInfo); override;
    procedure OnClientDisconnect(Client: TClientInfo); override;
    procedure OnClientSetName(Client: TClientInfo); override;
    procedure OnAddGroup(Group: TGroupInfo); override;
    procedure OnRemoveGroup(Group: TGroupInfo); override;
    procedure OnAddItem(Item: TGroupItemInfo); override;
    procedure OnRemoveItem(Item: TGroupItemInfo); override;
  public
  published
  end;

implementation
uses
  prOpcError, Windows, MainUnit, ComCtrls, TypInfo;

{ TDemo8 }

procedure TDemo8.OnClientConnect(Client: TClientInfo);
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

procedure TDemo8.OnClientDisconnect(Client: TClientInfo);
begin
  {Code here will execute whenever a client disconnects}
  MainForm.TreeView.Items.Delete(TTreeNode(Client.Data));
  LogMsg('Client disconnected Ok')
end;

procedure TDemo8.OnClientSetName(Client: TClientInfo);
begin
  {Code here will execute whenever a client calls IOpcCommon.SetClientName}
  TTreeNode(Client.Data).Text:=
    Format('Client %s',
      [Client.ClientName]);
  LogMsg('Client SetName Ok')
end;

procedure TDemo8.OnAddGroup(Group: TGroupInfo);
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

procedure TDemo8.OnRemoveGroup(Group: TGroupInfo);
begin
  {Code here will execute whenever a client removes a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Group.Data));
  LogMsg(Format('Remove Group %s ok', [Group.Name]))
end;

procedure TDemo8.OnAddItem(Item: TGroupItemInfo);
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

procedure TDemo8.OnRemoveItem(Item: TGroupItemInfo);
begin
  {Code here will execute whenever a client removes an item from a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Item.Data));
  LogMsg(Format('Remove Item %s ok', [Item.ItemID]))
end;

procedure TDemo8.LogMsg(const Msg: string);
begin
  if Assigned(MainForm) then
    MainForm.DebugLog.Items.Add(Msg)
end;

const
  ServerGuid: TGUID = '{11196550-1348-11D5-944D-00C0F023FA1C}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Rtti proxy demo';
  ServerVendor = 'Production Robots Eng. Ltd.';

function TDemo8.IsRecursive: Boolean;
begin
  Result:= false
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo8.Create)
end.

