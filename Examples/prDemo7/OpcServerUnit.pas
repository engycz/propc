{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{                                                            }
{ mailto engycz@gmail.com                                    }
{ http://code.google.com/p/propc/                            }
{ original source mailto: prOpcKit@prel.co.uk                }
{ original source http://www.prel.co.uk                      }
{------------------------------------------------------------}
unit OpcServerUnit;

interface

uses
  SysUtils, Windows, Classes, prOpcRttiServer, prOpcServer, prOpcTypes;

type
  TDemo7 = class(TRttiItemServer)
  private
    FClientCount: Integer;
    procedure UpdateClientCount;
    function GetTickCount: Integer;
    function GetTimeOfDay: String;
  protected
    procedure OnClientConnect(Client: TClientInfo); override;
    procedure OnClientDisconnect(Client: TClientInfo); override;
  public
  published
    {declare your Opc Items in here}
    property TickCount: Integer read GetTickCount;
    property TimeOfDay: String read GetTimeOfDay;
  end;

var
  MsgWnd: hWnd;

implementation
uses
  prOpcError;

{ TDemo1 }


const
  ServerGuid: TGUID = '{709E5F11-1336-11D5-944D-00C0F023FA1C}';
  ServerVersion = 1;
  ServerDesc = 'Simple non-VCL Demo';
  ServerVendor = 'Production Robots Eng Ltd.';

{ TDemo1 }

function TDemo7.GetTickCount: Integer;
begin
  Result:= Windows.GetTickCount
end;

function TDemo7.GetTimeOfDay: String;
begin
  Result:= TimeToStr(Now)
end;

procedure TDemo7.OnClientConnect(Client: TClientInfo);
begin
  Inc(FClientCount);
  UpdateClientCount
end;

procedure TDemo7.OnClientDisconnect(Client: TClientInfo);
begin
  Dec(FClientCount);
  UpdateClientCount
end;

procedure TDemo7.UpdateClientCount;
begin
  if MsgWnd <> 0 then
    SetWindowText(MsgWnd, PChar(Format('There are %d client(s) connected', [FClientCount])))
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo7.Create)
end.

