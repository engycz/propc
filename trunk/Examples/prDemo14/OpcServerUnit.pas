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
  SysUtils, Classes, prOpcServer, prOpcTypes, ComCtrls;

type
  TDemo14 = class(TOpcItemServer)
  private
    procedure LogMsg(const Msg: string);
  protected
    function Options: TServerOptions; override;
    function GetItemInfo(const ItemID: String;
                        var AccessPath: String;
                        var AccessRights: TAccessRights): Integer; override;
    function GetItemValue(ItemHandle: TItemHandle;
                          var Quality: Word): OleVariant; override;
    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); override;
    procedure ListItemIDs(List: TItemIDList); override;

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
  prOpcError, Windows, MainUnit, TypInfo;

{ TDemo5 }

procedure TDemo14.OnClientConnect(Client: TClientInfo);
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

procedure TDemo14.OnClientDisconnect(Client: TClientInfo);
begin
  {Code here will execute whenever a client disconnects}
  MainForm.TreeView.Items.Delete(TTreeNode(Client.Data));
  if ClientCount = 0 then
    MainForm.ExitServerButton.Enabled:= true;
  LogMsg('Client disconnected Ok');
end;

procedure TDemo14.OnClientSetName(Client: TClientInfo);
begin
  {Code here will execute whenever a client calls IOpcCommon.SetClientName}
  TTreeNode(Client.Data).Text:=
    Format('Client: %s', [Client.ClientName]);
  LogMsg('Client SetName Ok')
end;

procedure TDemo14.OnAddGroup(Group: TGroupInfo);
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

procedure TDemo14.OnRemoveGroup(Group: TGroupInfo);
begin
  {Code here will execute whenever a client removes a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Group.Data));
  LogMsg(Format('Remove Group %s ok', [Group.Name]))
end;

procedure TDemo14.OnAddItem(Item: TGroupItemInfo);
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

procedure TDemo14.OnRemoveItem(Item: TGroupItemInfo);
begin
  {Code here will execute whenever a client removes an item from a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Item.Data));
  LogMsg(Format('Remove Item %s ok', [Item.ItemID]))
end;

procedure TDemo14.LogMsg(const Msg: string);
begin
  if Assigned(MainForm) then
    MainForm.DebugLog.Items.Add(Msg)
end;

const
  ServerGuid: TGUID = '{8C213493-742D-419D-BC0C-04678B60354C}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Hierarchical Browsing Demo';
  ServerVendor = 'Production Robots Eng. Ltd.';

function TDemo14.GetItemInfo(const ItemID: String; var AccessPath: String;
  var AccessRights: TAccessRights): Integer;

var
  FoundNode: TTreeNode;

  function AddToPath(const Path, Str: string): string;
  begin
    if Path = '' then
      Result:= Str
    else
      Result:= Path + PathDelimiter + Str
  end;

  procedure ScanLevel(const Path: string; Node: TTreeNode);
  begin
    repeat
      if Node.HasChildren then {recurse}
      begin
        ScanLevel(AddToPath(Path, Node.Text), Node.GetFirstChild)
      end else
      begin {add name}
        if AddToPath(Path, TItemInfo(Node.Data).Name) = ItemId then
        begin
          FoundNode:= Node;
          break
        end
      end;
      Node:= Node.GetNextSibling
    until (Node = nil) or Assigned(FoundNode);
  end;

begin
  FoundNode:= nil;
  ScanLevel('', MainForm.Namespace.Items.GetFirstNode);
  if FoundNode = nil then
    raise EOpcError.Create(OPC_E_INVALIDITEMID);
  Result:= Integer(FoundNode)
end;

function TDemo14.GetItemValue(ItemHandle: TItemHandle;
  var Quality: Word): OleVariant;
var
  tn: TTreeNode;
begin
  tn:= TTreeNode(ItemHandle);
  Result:= TItemInfo(tn.Data).Buf
end;

procedure TDemo14.SetItemValue(ItemHandle: TItemHandle;
  const Value: OleVariant);
var
  tn: TTreeNode;
begin
  tn:= TTreeNode(ItemHandle);
  with TItemInfo(tn.Data) do
  begin
    Buf:= Value;
    case _Type of
      varInteger: tn.Text:= Name + '=' + IntToStr(Value);
      varDouble: tn.Text:= Name + '=' + FloatToStrF(Value, ffFixed, 6, 2);
      varOleStr: tn.Text:= Name + '=' + Value
    end
  end
end;

procedure TDemo14.ListItemIDs(List: TItemIDList);

procedure AddNewBranch(Node: TTreeNode; List: TItemIDList);
begin
  while Node <> nil do
  begin
    if Node.HasChildren then
    begin
      AddNewBranch(Node.GetFirstChild, List.NewBranch(Node.Text))
    end else
    begin
      with TItemInfo(Node.Data) do
        List.AddItemID(Name, AllAccess, _Type);
    end;
    Node:= Node.GetNextSibling
  end
end;

begin
  AddNewBranch(MainForm.Namespace.Items.GetFirstNode, List)
end;


function TDemo14.Options: TServerOptions;
begin
  Result:= [soHierarchicalBrowsing]
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo14.Create)
end.

