{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{                                                            }
{ mailto engycz@gmail.com                                    }
{ http://code.google.com/p/propc/                            }
{ original source mailto: prOpcKit@prel.co.uk                }
{ original source http://www.prel.co.uk                      }
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
