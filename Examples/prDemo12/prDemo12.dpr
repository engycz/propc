program prDemo12;

uses
  SvcMgr,
  MainUnit in 'MainUnit.pas' {Demo12Service: TService},
  OpcServerUnit in 'OpcServerUnit.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TDemo12Service, Demo12Service);
  Application.Run;
end.
