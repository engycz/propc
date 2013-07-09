{------------------------------------------------------------}
{The MIT License (MIT)

 prOpc Toolkit
 Copyright (c) 2000, 2001 Production Robots Engineering Ltd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.}
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
