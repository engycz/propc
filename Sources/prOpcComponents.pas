unit prOpcComponents;

interface

uses Classes, prOpcClient, prOpcBrowser;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('OPC', [TOpcSimpleClient, TOpcBrowser, TOpcPropertyView]);
end;

end.
