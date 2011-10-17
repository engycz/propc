{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit OpcServerUnit;

interface

uses
  SysUtils, Classes, Windows, prOpcRttiServer, prOpcServer, prOpcTypes;

type
  TDemo12 = class(TRttiItemServer)
  private
    FReadWriteItem: Integer;
    function GetTimeOfDay: String;
    function GetTickCount: Cardinal;
    function GetServiceStatus: Integer;
  protected
  public
  published
    {declare your Opc Items in here}
    property TimeOfDay: String read GetTimeOfDay;
    property TickCount: Cardinal read GetTickCount;
    property ReadWriteItem: Integer read FReadWriteItem write FReadWriteItem;
    property ServiceStatus: Integer read GetServiceStatus;
  end;

implementation
uses
  MainUnit, prOpcError;

{ TDemo12 }

function TDemo12.GetServiceStatus: Integer;
begin
  if Assigned(Demo12Service) then
    Result:= Integer(Demo12Service.Status)
  else
    Result:= 0
end;

function TDemo12.GetTickCount: Cardinal;
begin
  Result:= Windows.GetTickCount
end;

function TDemo12.GetTimeOfDay: String;
begin
  Result:= DateToStr(Now)
end;


const
  ServerGuid: TGUID = '{C9161755-AE8A-4390-BD1E-F56EBFD29DE4}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit Demo12 - Service Application';
  ServerVendor = 'Production Robots Engineering Ltd';

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo12.Create)
end.
