library prDemo6;

uses
  ComServ,
  OpcServerUnit in 'OpcServerUnit.pas',
  prDemo6_TLB in 'prDemo6_TLB.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
