program prDemoComplianceTest;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {Form1},
  OpcServerUnit in 'OpcServerUnit.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
