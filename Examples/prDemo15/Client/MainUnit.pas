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
{Notes:
 This demo is intended to demonstrate the use of EnumeratedEU. It requres that
 the deadband and EnumeratedEU server prDemo15, be built and registered.

Note that the 'PercentDeadband' property for group 0 is set to 10}

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
    LogLabel: TLabel;
    TrackBar: TTrackBar;
    ListBox: TListBox;
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
    procedure TrackBarChange(Sender: TObject);
    procedure ListBoxClick(Sender: TObject);
  private
    FProcessingDataChange: Boolean;
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
var
  LowRange, HighRange: Double;
  Group: TOpcGroup;
begin
  Client.Active:= true;
  {Item 0 is the track bar. We can ask the server for the
  nominal range of this item}
  Group:= Client.Groups[0];
  if not Group.ItemAnalogRange(0, LowRange, HighRange) then
    raise Exception.Create('Item 0 does not supply range information');
  {we can now apply this range to the trackbar}
  Trackbar.Min:= Round(LowRange);
  Trackbar.Max:= Round(HighRange);

  {Item 1 is the list box. This is an integer type which provides a name
  for each possible value it might take. We can ask the server for these
  names}
  if not Group.ItemEnumeratedNames(1, ListBox.Items) then
    raise Exception.Create('Item 1 does not supply enumerated names');

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

procedure TMainForm.ClientServerShutdown(Sender: TObject);
begin
  MsgLog.Items.Add('Server shutdown req: ' + Client.ShutdownReason)
end;

procedure TMainForm.ClientGroups0DataChange(Sender: TOpcGroup;
  ItemIndex: Integer; const NewValue: Variant; NewQuality: Word;
  NewTimestamp: TDateTime);
begin
  FProcessingDataChange:= true; {don't allow change handlers to send
                                a new value to the server}
  try
    case ItemIndex of
      0: TrackBar.Position:= NewValue;
      1: ListBox.ItemIndex:= NewValue;
    end
  finally
    FProcessingDataChange:= false
  end
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Client.Active:= false
end;

procedure TMainForm.TrackBarChange(Sender: TObject);
begin
  if Client.Active and not FProcessingDataChange then
    Client.Groups[0].ItemValue[0]:= TrackBar.Position
end;

procedure TMainForm.ListBoxClick(Sender: TObject);
begin
  if Client.Active and not FProcessingDataChange then
    Client.Groups[0].ItemValue[1]:= ListBox.ItemIndex
end;

end.
