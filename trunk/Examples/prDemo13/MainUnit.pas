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
  ComCtrls, ExtCtrls, StdCtrls;

type
  TMainForm = class(TForm)
    DebugLog: TListBox;
    Panel1: TPanel;
    ShutdownClientsButton: TButton;
    ExitServerButton: TButton;
    procedure ExitServerButtonClick(Sender: TObject);
    procedure ShutdownClientsButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
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
  OpcItemServer.ShutdownRequest('prDemo13: server request shutdown')
end;

end.
