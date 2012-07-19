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
  SysUtils, Classes, prOpcRttiServer, prOpcServer, prOpcTypes;

type
  TDemo6 = class(TRttiItemServer)
  private
    function GetTickCount: Integer;
    function GetTimeOfDay: string;
  protected
  public
  published
    property TickCount: Integer read GetTickCount;
    property TimeOfDay: string read GetTimeOfDay;
  end;

implementation
uses
  prOpcError, Windows;


{ TDemo6 }

const
  ServerGuid: TGUID = '{E25AA401-121D-11D5-944C-00C0F023FA1C}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - In process demo';
  ServerVendor = 'Production Robots Eng. Ltd.';


{ TDemo6 }

function TDemo6.GetTickCount: Integer;
begin
  Result:= Windows.GetTickCount
end;

function TDemo6.GetTimeOfDay: string;
begin
  Result:= TimeToStr(Now)
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo6.Create)
end.
