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
unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  prOpcClient, StdCtrls, prOpcTypes, ComCtrls, prOpcUtils,
  prOpcServerSelect;

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

procedure TMainForm.FormCreate(Sender: TObject);
begin
  UpdateButtons
end;

procedure TMainForm.ConnectButtonClick(Sender: TObject);
var
  Group: TOpcGroup;
  i: Integer;
  j: Integer;
  SelectDlg: TServerSelectDlg;
begin
  SelectDlg:= TServerSelectDlg.Create(Self);
  try
    if SelectDlg.Execute(Client) then
    begin
      Group:= Client.Groups[0];
      Group.Items.Clear;
      Client.Connect(true);
      Client.GetAllItems(Group.Items);
      ItemList.Items.BeginUpdate;
      ItemList.Items.Clear;
      try
        for i:= 0 to Group.Items.Count - 1 do
        with ItemList.Items.Add do
        begin
          Caption:= Group.Items[i];
          for j:= 0 to 2 do
            Subitems.Add('')
        end
      finally
        ItemList.Items.EndUpdate
      end;
      Client.ConnectGroups
    end
  finally
    SelectDlg.Free
  end
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
    SubItems[0]:= NewValue;
    SubItems[1]:= QualityToStr(NewQuality);
    SubItems[2]:= TimeToStr(NewTimestamp);
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
