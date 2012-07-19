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
  {$M+}  {must be compiled with RTTI on}
  TObj = class
  private
    FBool: Boolean;
    FFloat: Double;
    FInt: Integer;
    FStr: String;
  published
    property Str: String read FStr write FStr;
    property Int: Integer read FInt write FInt;
    property Bool: Boolean read FBool write FBool;
    property Float: Double read FFloat write FFloat;
  end;
  {$M-}

  TDemo9 = class(TRttiItemServer)
  private
    FObj: array[0..3] of TObj;
    procedure LogMsg(const Msg: string);
    function GetTickCount: Integer;
    function GetTimeOfDay: String;
  protected
    function Options: TServerOptions; override;
    procedure OnClientConnect(Client: TClientInfo); override;
    procedure OnClientDisconnect(Client: TClientInfo); override;
    procedure OnClientSetName(Client: TClientInfo); override;
    procedure OnAddGroup(Group: TGroupInfo); override;
    procedure OnRemoveGroup(Group: TGroupInfo); override;
    procedure OnAddItem(Item: TGroupItemInfo); override;
    procedure OnRemoveItem(Item: TGroupItemInfo); override;
    procedure LoadRttiItems(Proxy: TObject); override;
  public
    destructor Destroy; override;
  published
    property TickCount: Integer read GetTickCount;
    property TimeOfDay: String read GetTimeOfDay;
    property Obj0: TObj read FObj[0];
    property Obj1: TObj read FObj[1];
    property Obj2: TObj read FObj[2];
    property Obj3: TObj read FObj[3];
  end;

implementation
uses
  prOpcError, Windows, MainUnit, ComCtrls;

{ TDemo9 }

function TDemo9.Options: TServerOptions;
begin
  Result:= Options + [soHierarchicalBrowsing]
end;

procedure TDemo9.OnClientConnect(Client: TClientInfo);
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

procedure TDemo9.OnClientDisconnect(Client: TClientInfo);
begin
  {Code here will execute whenever a client disconnects}
  MainForm.TreeView.Items.Delete(TTreeNode(Client.Data));
  LogMsg('Client disconnected Ok')
end;

procedure TDemo9.OnClientSetName(Client: TClientInfo);
begin
  {Code here will execute whenever a client calls IOpcCommon.SetClientName}
  TTreeNode(Client.Data).Text:=
    Format('Client %s',
      [Client.ClientName]);
  LogMsg('Client SetName Ok')
end;

procedure TDemo9.OnAddGroup(Group: TGroupInfo);
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

procedure TDemo9.OnRemoveGroup(Group: TGroupInfo);
begin
  {Code here will execute whenever a client removes a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Group.Data));
  LogMsg(Format('Remove Group %s ok', [Group.Name]))
end;

procedure TDemo9.OnAddItem(Item: TGroupItemInfo);
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

procedure TDemo9.OnRemoveItem(Item: TGroupItemInfo);
begin
  {Code here will execute whenever a client removes an item from a group}
  MainForm.TreeView.Items.Delete(TTreeNode(Item.Data));
  LogMsg(Format('Remove Item %s ok', [Item.ItemID]))
end;

procedure TDemo9.LogMsg(const Msg: string);
begin
  if Assigned(MainForm) then
    MainForm.DebugLog.Items.Add(Msg)
end;

const
  ServerGuid: TGUID = '{F4EB7AA0-1348-11D5-944D-00C0F023FA1C}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Recursive Rtti demo';
  ServerVendor = 'Production Robots Eng. Ltd.';

function TDemo9.GetTickCount: Integer;
begin
  Result:= Windows.GetTickCount
end;

function TDemo9.GetTimeOfDay: String;
begin
  Result:= TimeToStr(Now)
end;

destructor TDemo9.Destroy;
var
  i: Integer;
begin
  for i:= 0 to 3 do
    FObj[i].Free;
  inherited Destroy
end;

procedure TDemo9.LoadRttiItems(Proxy: TObject);
var
  i: Integer;
begin
  for i:= 0 to 3 do
    FObj[i]:= TObj.Create;
  inherited LoadRttiItems(Proxy)
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo9.Create)
end.

