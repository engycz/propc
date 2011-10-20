{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit prOpcBrowser;
{$I prOpcCompilerDirectives.inc}
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, prOpcClient, prOpcUtils, prOpcTypes, prOpcDa;

type
  TOpcNode = class(TTreeNode)
  public
    ItemID: String;
  end;

  TOpcPropertyView = class(TCustomListView)
  private
    procedure Populate(Props: TItemProperties);
    function GetColWidth(i: Integer): Integer;
    procedure SetColWidth(i: Integer; Value: Integer);
  public
    constructor Create(AComponent: TComponent); override;
    procedure ShowProperties(Client: TOpcSimpleClient; const ItemID: string);
  published
    property ColWidthID: Integer index 0 read GetColWidth write SetColWidth;
    property ColWidthDesc: Integer index 1 read GetColWidth write SetColWidth;
    property ColWidthDatatype: Integer index 2 read GetColWidth write SetColWidth;
    property ColWidthValue: Integer index 3 read GetColWidth write SetColWidth;
    property Align;
    property Anchors;
    property BiDiMode;
    property BorderStyle;
    property BorderWidth;
    property Color;
    property Constraints;
    property Ctl3D;
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property Font;
    property FlatScrollBars;
    property FullDrag;
    property GridLines;
    property HideSelection;
    property HotTrack;
    property HotTrackStyles;
    property HoverTime;
    property RowSelect;
    property ParentBiDiMode;
    property ParentColor default False;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property ShowHint;
    property SortType;
    property TabOrder;
    property TabStop default True;
    property Visible;
    property OnChange;
    property OnChanging;
    property OnClick;
    property OnColumnClick;
    property OnColumnDragged;
    property OnColumnRightClick;
    property OnContextPopup;
    property OnDataStateChange;
    property OnDblClick;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnDragDrop;
    property OnDragOver;
    property OnInfoTip;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnResize;
    property OnSelectItem;
    property OnStartDock;
    property OnStartDrag;
  end;

  TOpcBrowser = class(TCustomTreeView)
  private
    FOpcClient: TOpcSimpleClient;
    FPropView: TOpcPropertyView;
    procedure WMOpcClientUpdate(var Msg: TMessage); message WM_OPCBROWSEUPDATE;
    procedure SetOpcClient(Value: TOpcSimpleClient);
  protected
    function CreateNode: TTreeNode; override;
    procedure Change(Node: TTreeNode); override;
    function BrowseNodeEvent(Parent: Pointer; const BrowseId, ItemId: string): Pointer;
  public
    function SelectedItemID: string;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure RefreshClient(aClient: TOpcSimpleClient);
    class function GetNodeItemID(Node: TTreeNode): string;
  published
    property OpcClient: TOpcSimpleClient read FOpcClient write SetOpcClient;
    property PropertyView: TOpcPropertyView read FPropView write FPropView;
    property Align;
    property Anchors;
    property AutoExpand;
    property BiDiMode;
    property BorderStyle;
    property BorderWidth;
    property ChangeDelay;
    property Color;
    property Ctl3D;
    property Constraints;
    property DragKind;
    property DragCursor;
    property DragMode;
    property Enabled;
    property Font;
    property HideSelection;
    property HotTrack;
    property Indent;
    property ParentBiDiMode;
    property ParentColor default False;
    property ParentCtl3D;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property RightClickSelect;
    property RowSelect;
    property ShowButtons;
    property ShowHint;
    property ShowLines;
    property ShowRoot;
    property SortType;
    property StateImages;
    property TabOrder;
    property TabStop default True;
    property ToolTips;
    property Visible;
    property OnChange;
    property OnChanging;
    property OnClick;
    property OnCollapsed;
    property OnCollapsing;
    property OnCompare;
    property OnContextPopup;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnExpanding;
    property OnExpanded;
    property OnGetImageIndex;
    property OnGetSelectedIndex;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnStartDock;
    property OnStartDrag;
    { Items must be published after OnGetImageIndex and OnGetSelectedIndex }
    property Items;
  end;

implementation
uses
  ComObj, ActiveX;

resourcestring
  SBadNodeType = 'This node does not contain Item ID';

{ TOpcBrowser }

procedure TOpcBrowser.Change(Node: TTreeNode);
begin
  if Assigned(FPropView) then
    FPropView.ShowProperties(FOpcClient,
      TOpcNode(Node).ItemId);
  inherited Change(Node)
end;

function TOpcBrowser.CreateNode: TTreeNode;
begin
  Result:= TOpcNode.Create(Items)
end;

function TOpcBrowser.BrowseNodeEvent(Parent: Pointer; const BrowseId,
  ItemId: string): Pointer;
begin
  Result:= Items.AddChild(Parent, BrowseID);
  TOpcNode(Result).ItemId:= ItemId;
end;

class function TOpcBrowser.GetNodeItemID(Node: TTreeNode): string;
begin
  if Assigned(Node) and (Node is TOpcNode) then
    Result:= TOpcNode(Node).ItemID
  else
    raise EOpcClient.CreateRes(@SBadNodeType)
end;

procedure TOpcBrowser.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if Operation = opRemove then
  begin
    if (AComponent = FOpcClient) then
      OpcClient:= nil
    else
    if (AComponent = FPropView) then
      FPropView:= nil
  end
end;

procedure TOpcBrowser.RefreshClient(aClient: TOpcSimpleClient);
begin
  {Don't do this at design time}
  if [csDesigning, csLoading] * ComponentState = [] then
  begin
    Items.Clear;
    if Assigned(aClient) then
      aClient.BrowseItems(BrowseNodeEvent);
  end
end;

function TOpcBrowser.SelectedItemID: string;
begin
  Result:= GetNodeItemID(Selected)
end;

procedure TOpcBrowser.SetOpcClient(Value: TOpcSimpleClient);
var
  Msg: TMessage;
begin
  if FOpcClient <> Value then
  begin
    Msg.Msg:= WM_OPCBROWSEUPDATE;
    Msg.WParam:= 0;
    Msg.LParam:= 0;
    Msg.Result:= 0;
    if Assigned(FOpcClient) then
      FOpcClient.Dispatch(Msg);
    if Assigned(Value) then
    begin
      Msg.WParam:= Integer(Self);
      Value.Dispatch(Msg)
    end;
    FOpcClient:= Value;
    RefreshClient(FOpcClient)
  end
end;

procedure TOpcBrowser.WMOpcClientUpdate(var Msg: TMessage);
var
  Client: TOpcSimpleClient;
begin
  Client:= TOpcSimpleClient(Msg.wParam);
  if Assigned(Client) then
    RefreshClient(Client);
  Msg.Result:= 1
end;

{ TOpcPropertyView }

type
  TPropCols = (pcID, pcDesc, pcType, pcValue);
const
  ColNames: array[TPropCols] of String = ('ID', 'Description', 'Type', 'Value');
  ColWidths: array[TPropCols] of Integer = (40, 200, 80, 200);

constructor TOpcPropertyView.Create(AComponent: TComponent);
var
  i: TPropCols;
begin
  inherited Create(AComponent);
  ReadOnly:= true;
  ViewStyle:= vsReport;
  Columns.BeginUpdate;
  try
    for i:= Low(TPropCols) to High(TPropCols) do
    with Columns.Add do
    begin
      Caption:= ColNames[i];
      Width:= ColWidths[i]
    end
  finally
    Columns.EndUpdate
  end
end;

function InterpretValue(ID: Integer; const Value: OleVariant): String;
begin
  case ID of
    1: Result:= DatatypeToStr(Value);
    3: Result:= QualityToStr(Value);
    4: Result:= DateTimeToStr(Value);
    5: Result:= AccessRightsToStr(Value);
  else
    Result:= Value
  end
end;

function TOpcPropertyView.GetColWidth(i: Integer): Integer;
begin
  with Columns[i] do
  if Autosize then
    Result:= 0
  else
    Result:= Width
end;

procedure TOpcPropertyView.Populate(Props: TItemProperties);
var
  i: Integer;
begin
  Items.BeginUpdate;
  try
    Items.Clear;
    for i:= 0 to Length(Props) - 1 do
    with Items.Add, Props[i] do
    begin
      Caption:= IntToStr(ID);
      SubItems.Add(Desc);
      SubItems.Add(DatatypeToStr(Datatype));
      SubItems.Add(InterpretValue(ID, Value))
    end
  finally
    Items.EndUpdate
  end
end;

procedure TOpcPropertyView.SetColWidth(i, Value: Integer);
begin
  with Columns[i] do
  begin
    if Value = 0 then
    begin
      AutoSize:= true
    end else
    begin
      AutoSize:= false;
      Width:= Value
    end
  end
end;

procedure TOpcPropertyView.ShowProperties(Client: TOpcSimpleClient;
  const ItemID: string);
var
  Props: TItemProperties;
begin
  if Assigned(Client) and Client.Active and (ItemId <> '') then
  begin
    Client.GetItemProperties(ItemId, Props);
    Populate(Props)
  end else
  begin
    Items.Clear
  end
end;

end.
