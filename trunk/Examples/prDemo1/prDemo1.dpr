program prDemo1;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {MainForm},
  OpcServerUnit in 'OpcServerUnit.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
