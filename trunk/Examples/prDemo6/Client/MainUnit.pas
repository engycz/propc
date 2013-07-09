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
  prOpcClient, StdCtrls, prOpcTypes, ComCtrls, prOpcUtils;

type
  TMainForm = class(TForm)
    Client: TOpcSimpleClient;
    ConnectButton: TButton;
    DisconnectButton: TButton;
    MsgLog: TListBox;
    Label1: TLabel;
    TickCountValue: TStaticText;
    TimeOfDayValue: TStaticText;
    Label2: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure ConnectButtonClick(Sender: TObject);
    procedure DisconnectButtonClick(Sender: TObject);
    procedure ClientConnect(Sender: TObject);
    procedure ClientDisconnect(Sender: TObject);
    procedure ClientGroups0DataChange(Sender: TOpcGroup;
      ItemIndex: Integer; const NewValue: Variant; NewQuality: Word;
      NewTimestamp: TDateTime);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    procedure UpdateButtons;
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.DFM}

procedure TMainForm.FormCreate(Sender: TObject);
begin
  UpdateButtons
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
  DisconnectButton.Enabled:= Client.Active
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

procedure TMainForm.ClientGroups0DataChange(Sender: TOpcGroup;
  ItemIndex: Integer; const NewValue: Variant; NewQuality: Word;
  NewTimestamp: TDateTime);
begin
  case ItemIndex of
    0: TimeOfDayValue.Caption:= NewValue;
    1: TickCountValue.Caption:= IntToStr(NewValue);
  end
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Client.Active:= false
end;

end.
