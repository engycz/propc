unit MainUnit;

interface
{release 1.15b. Fixed for delphi5}

uses
  Windows, Messages, SysUtils,
{$IFNDEF VER130}
  Variants,
{$ENDIF}
  Classes, Graphics, Controls, Forms,
  Dialogs, prOpcVSC, StdCtrls, ExtCtrls;

type
  TMainForm = class(TForm)
    StartButton: TButton;
    procedure StartButtonClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    { Private declarations }
  public
    VSC: Variant;
  end;

var
  MainForm: TMainForm;

implementation
{1. Make sure that prDemo5.exe is registered.
Run this program
Click Start
Resize the client form: Observe the server.
}


{$R *.dfm}

procedure TMainForm.StartButtonClick(Sender: TObject);
begin
  if VarIsEmpty(VSC) then
  begin
    VSC:= OpcClient('', 'prDemo5.TDemo5.1', 0);
    SetBounds(Left, Top, VSC.FormWidth, VSC.FormHeight);
    StartButton.Caption:= 'Stop'
  end else
  begin
    VSC:= Unassigned;
    StartButton.Caption:= 'Start'
  end
end;

procedure TMainForm.FormResize(Sender: TObject);
begin
  if not VarIsEmpty(VSC) then
  begin
    VSC.FormWidth:= Width;
    VSC.FormHeight:= Height
  end
end;

end.
