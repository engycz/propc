unit OpcServerUnit;

interface

uses
  SysUtils, Classes, prOpcRttiServer, prOpcServer, prOpcTypes;

type
  TWizzyServer = class(TRttiItemServer)
  private
    function GetFormHeight: Integer;
    function GetFormWidth: Integer;
    procedure SetFormHeight(const Value: Integer);
    procedure SetFormWidth(const Value: Integer);
  protected
  public
  published
    {declare your Opc Items in here}
    property FormHeight: Integer read GetFormHeight write SetFormHeight;
    property FormWidth: Integer read GetFormWidth write SetFormWidth;
  end;

implementation
uses
  MainUnit,
  prOpcError;

{ TWizzyServer }

const
  ServerGuid: TGUID = '{F1B1B47F-BC2C-4E88-A579-F24125855B9B}';
  ServerVersion = 1;
  ServerDesc = 'Simple but Wizzy Server';
  ServerVendor = 'Wizzy Vendor';

{ TWizzyServer }

function TWizzyServer.GetFormHeight: Integer;
begin
  Result:= MainForm.Height
end;

function TWizzyServer.GetFormWidth: Integer;
begin
  Result:= MainForm.Width
end;

procedure TWizzyServer.SetFormHeight(const Value: Integer);
begin
  MainForm.Height:= Value
end;

procedure TWizzyServer.SetFormWidth(const Value: Integer);
begin
  MainForm.Width:= Value
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TWizzyServer.Create)
end.
