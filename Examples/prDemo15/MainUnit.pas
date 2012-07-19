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
{prDemo15 - PercentageDeadband and Enumerated Types}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ExtCtrls, StdCtrls;

type
  TMainForm = class(TForm)
    ExitServerButton: TButton;
    ShutdownClientsButton: TButton;
    TrackBar: TTrackBar;
    ListBox: TComboBox;
    procedure ExitServerButtonClick(Sender: TObject);
    procedure ShutdownClientsButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    procedure ShowDeadband(LastVal: Integer; Deadband: Single);
  end;

var
  MainForm: TMainForm;

implementation
uses
  prOpcServer, prOpcTypes;

{$R *.DFM}

procedure TMainForm.ExitServerButtonClick(Sender: TObject);
begin
  Close
end;

procedure TMainForm.ShutdownClientsButtonClick(Sender: TObject);
begin
  OpcItemServer.ShutdownRequest('prDemo15: server request shutdown')
end;

procedure TMainForm.ShowDeadband(LastVal: Integer; Deadband: Single);
var
  IntDb: Integer;
begin
  IntDb:= Round(Deadband);
  TrackBar.SelStart:= LastVal - IntDb;
  TrackBar.SelEnd:= LastVal + IntDb
end;

end.

