{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  prOpcClient, StdCtrls, prOpcTypes, ComCtrls, prOpcUtils;

type
  TMainForm = class(TForm)
    Client: TOpcSimpleClient;
    ConnectButton: TButton;
    DisconnectButton: TButton;
    MsgLog: TListBox;
    ItemList: TListView;
    GroupLabel: TLabel;
    LogLabel: TLabel;
    BrowseButton: TButton;
    WriteButton: TButton;
    procedure FormCreate(Sender: TObject);
    procedure ConnectButtonClick(Sender: TObject);
    procedure DisconnectButtonClick(Sender: TObject);
    procedure ClientConnect(Sender: TObject);
    procedure ClientDisconnect(Sender: TObject);
    procedure ClientServerShutdown(Sender: TObject);
    procedure ClientGroups0DataChange(Sender: TOpcGroup;
      ItemIndex: Integer; const NewValue: Variant; NewQuality: Word;
      NewTimestamp: TDateTime);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BrowseButtonClick(Sender: TObject);
    procedure ItemListChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure WriteButtonClick(Sender: TObject);
  private
    procedure UpdateButtons;
    procedure UpdateWriteButton;
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

uses BrowseUnit;

{$R *.DFM}

const
  SubItemValue = 0;
  SubItemQuality = 1;
  SubItemTimestamp = 2;

procedure TMainForm.FormCreate(Sender: TObject);
var
  Group: TOpcGroup;
  i: Integer;
  j: Integer;
begin
  UpdateButtons;
  Group:= Client.Groups[0];
  for i:= 0 to Group.Items.Count - 1 do
  with ItemList.Items.Add do
  begin
    Caption:= Group.Items[i];
    for j:= 0 to 2 do
      Subitems.Add('')
  end
end;

procedure TMainForm.ConnectButtonClick(Sender: TObject);
begin
  Client.Active:= true
end;

procedure TMainForm.DisconnectButtonClick(Sender: TObject);
begin
  Client.Active:= false
end;

procedure TMainForm.UpdateButtons;
begin
  ConnectButton.Enabled:= not Client.Active;
  DisconnectButton.Enabled:= Client.Active;
  BrowseButton.Enabled:= Client.Active;
  UpdateWriteButton
end;

procedure TMainForm.ClientConnect(Sender: TObject);
begin
  MsgLog.Items.Add('Connected: ' + DateTimeToStr(Now));
  UpdateButtons
end;

procedure TMainForm.ClientDisconnect(Sender: TObject);
begin
  MsgLog.Items.Add('Disconnected: ' + DateTimeToStr(Now));
  UpdateButtons
end;

procedure TMainForm.ClientServerShutdown(Sender: TObject);
begin
  MsgLog.Items.Add('Server shutdown req: ' + Client.ShutdownReason)
end;

procedure TMainForm.ClientGroups0DataChange(Sender: TOpcGroup;
  ItemIndex: Integer; const NewValue: Variant; NewQuality: Word;
  NewTimestamp: TDateTime);
begin
  with ItemList.Items[ItemIndex] do
  begin
    SubItems[SubItemValue]:= NewValue;
    SubItems[SubItemQuality]:= QualityToStr(NewQuality);
    SubItems[SubItemTimestamp]:= TimeToStr(NewTimestamp);
  end
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Client.Active:= false
end;

procedure TMainForm.BrowseButtonClick(Sender: TObject);
begin
  BrowseDlg.Execute(Client)
end;

procedure TMainForm.UpdateWriteButton;
var
  Sel: TListItem;
begin
  Sel:= ItemList.Selected;
  WriteButton.Enabled:= Client.Active and
    (Sel <> nil) and
    (iaWrite in Client.Groups[0].ItemAccessRights[Sel.Index])
end;

procedure TMainForm.ItemListChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
  UpdateWriteButton
end;

procedure TMainForm.WriteButtonClick(Sender: TObject);
var
  Val: String;
  i: Integer;
  Sel: TListItem;
  Group: TOpcGroup;
begin
  Sel:= ItemList.Selected;
  if Assigned(Sel) then
  begin
    i:= Sel.Index;
    Group:= Client.Groups[0];
    Val:= Group.ItemValue[i];
    if InputQuery('Sync Write',
      Format('Enter new value for %s', [Group.Items[i]]), Val) then
      Group.SyncWriteItem(i, Val)
  end
end;

end.
