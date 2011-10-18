unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Math, ExtCtrls, AppEvnts, Spin,
  ComCtrls, OpcServerUnit, ActiveX, prOpcItems;

type
  TForm1 = class(TForm)
    Label1: TLabel;
    StatusBar1: TStatusBar;
    Timer1: TTimer;
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }                           
    Inited : Boolean;
    OldL   : LongWord;
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  StatusBar1.Panels[0].Text := IntToStr(AllocMemCount)+' / '+IntToStr(AllocMemSize);
end;

end.

