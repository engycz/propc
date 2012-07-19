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
  ComCtrls, ExtCtrls, StdCtrls;

type
  TMainForm = class(TForm)
    Panel1: TPanel;
    ShutdownClientsButton: TButton;
    ExitServerButton: TButton;
    UpdateButton: TButton;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    procedure ExitServerButtonClick(Sender: TObject);
    procedure ShutdownClientsButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure UpdateButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation
uses
  prOpcServer, prOpcTypes, OpcServerUnit;

{$R *.DFM}

procedure TMainForm.ExitServerButtonClick(Sender: TObject);
begin
  Close
end;

procedure TMainForm.ShutdownClientsButtonClick(Sender: TObject);
begin
  OpcItemServer.ShutdownRequest('prDemo10: server request shutdown')
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  with TDemo10(OpcItemServer) do
  begin
    Edit1.Text:= WideStringData[0];
    Edit2.Text:= WideStringData[1];
    Edit3.Text:= WideStringData[2];
    Edit4.Text:= WideStringData[3];
    Edit5.Text:= WideStringData[4]
  end
end;

procedure TMainForm.UpdateButtonClick(Sender: TObject);
begin
  with TDemo10(OpcItemServer) do
  begin
    WideStringData[0]:= Edit1.Text;
    WideStringData[1]:= Edit2.Text;
    WideStringData[2]:= Edit3.Text;
    WideStringData[3]:= Edit4.Text;
    WideStringData[4]:= Edit5.Text
  end
end;

end.

